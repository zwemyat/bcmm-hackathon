<?php

namespace App\Console\Commands;

use App\Mail\ExpiryReminderDigest;
use App\Models\LicenseContract;
use App\Models\MailDigestSend;
use App\Models\MailSetting;
use App\Models\NotificationSetting;
use App\Models\Subscription;
use App\Models\User;
use Carbon\Carbon;
use Illuminate\Console\Command;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Facades\Mail;

class CheckExpirations extends Command
{
    protected $signature = 'app:check-expirations {--force : Re-send digests even if today\'s ledger row already exists}';

    protected $description = 'Mark overdue subscriptions Expired and email staggered renewal-reminder digests for each (module × selected day-mark) bucket. Idempotent within a single day — a successful send writes a mail_digest_sends row that blocks re-sends; pass --force to override.';

    /** Modules eligible for staggered reminders. */
    private const MODULES = [
        'subscriptions' => [
            'label' => 'Subscription',
            'flip_pending_on_send' => true,
        ],
        'licenses_contracts' => [
            'label' => 'License & Contract',
            'flip_pending_on_send' => false,
        ],
    ];

    public function handle(): int
    {
        $today = Carbon::today();

        // Honor DB-stored SMTP credentials when the admin enabled them on the
        // Mail Settings page. Without this, scheduled digests would silently
        // send via the .env mailer config regardless of UI choice.
        MailSetting::current()->applyRuntimeConfig();

        $this->markOverdueSubscriptions($today);

        $totalBatches = 0;
        foreach (array_keys(self::MODULES) as $moduleKey) {
            $totalBatches += $this->sendStaggeredFor($moduleKey, $today);
        }

        $this->info("Done. Sent {$totalBatches} digest batch(es) in total.");

        return self::SUCCESS;
    }

    private function markOverdueSubscriptions(Carbon $today): void
    {
        $expired = Subscription::where('status', 'Active')
            ->where('renewal_status', '!=', 'Renewed')
            ->whereDate('expire_date', '<', $today)
            ->get();

        foreach ($expired as $subscription) {
            $subscription->renewal_status = 'Expired';
            $subscription->saveQuietly();
        }

        $this->info("Marked {$expired->count()} subscription(s) as Expired.");
    }

    private function sendStaggeredFor(string $moduleKey, Carbon $today): int
    {
        $setting = NotificationSetting::query()->where('module', $moduleKey)->first();
        if (! $setting || ! $setting->enabled) {
            $this->info("[{$moduleKey}] notifications disabled — skipped.");
            return 0;
        }

        $days = $setting->selectedDays(); // descending unique ints
        if (empty($days)) {
            $this->info("[{$moduleKey}] no day-marks selected — skipped.");
            return 0;
        }

        $recipients = $setting->recipientsArray();
        if (empty($recipients)) {
            $recipients = User::where('role', 'admin')->pluck('email')->filter()->values()->toArray();
        }
        if (empty($recipients)) {
            $this->warn("[{$moduleKey}] no recipients (no admin users with emails) — skipped.");
            return 0;
        }

        $cfg = self::MODULES[$moduleKey];
        $batches = 0;
        $force = (bool) $this->option('force');

        foreach ($days as $d) {
            // Idempotency guard — skip a (module, day_mark, today) that already
            // succeeded. Manual re-runs can bypass with --force.
            if (! $force && MailDigestSend::query()
                ->where('module', $moduleKey)
                ->where('day_mark', $d)
                ->where('sent_on', $today->toDateString())
                ->exists()
            ) {
                $this->info("[{$moduleKey}] {$d}-day digest already sent today — skipped (use --force to resend).");
                continue;
            }

            $target = $today->copy()->addDays($d);
            $rows = $this->baseQuery($moduleKey)
                ->whereDate('expire_date', $target)
                ->orderBy('expire_date')
                ->get();

            if ($rows->isEmpty()) {
                continue;
            }

            // For subscriptions, flip renewal_status to Pending on send so the
            // UI shows them as "Pending renewal" until acted on.
            if ($cfg['flip_pending_on_send']) {
                foreach ($rows as $s) {
                    if ($s->renewal_status !== 'Pending') {
                        $s->renewal_status = 'Pending';
                        $s->saveQuietly();
                    }
                }
            }

            try {
                Mail::to($recipients)->send(new ExpiryReminderDigest(
                    moduleKey:   $moduleKey,
                    moduleLabel: $cfg['label'],
                    daysAhead:   $d,
                    records:     $rows,
                ));

                // Record the successful send so a later run (manual or scheduler
                // re-trigger) on the same day doesn't double-send. updateOrCreate
                // because --force can re-send and we want the latest counts.
                MailDigestSend::updateOrCreate(
                    [
                        'module'   => $moduleKey,
                        'day_mark' => $d,
                        'sent_on'  => $today->toDateString(),
                    ],
                    [
                        'records_count'    => $rows->count(),
                        'recipients_count' => count($recipients),
                        'sent_at'          => now(),
                    ],
                );

                $batches++;
                $this->info("[{$moduleKey}] sent {$d}-day digest with {$rows->count()} record(s) to " . count($recipients) . ' recipient(s).');
            } catch (\Throwable $e) {
                $this->error("[{$moduleKey}] failed to send {$d}-day digest: {$e->getMessage()}");
            }
        }

        if ($batches === 0) {
            $this->info("[{$moduleKey}] no records matched any selected day-mark today.");
        }

        return $batches;
    }

    private function baseQuery(string $moduleKey): Builder
    {
        return match ($moduleKey) {
            'subscriptions' => Subscription::query()
                ->where('status', 'Active')
                ->where('renewal_status', '!=', 'Renewed'),
            'licenses_contracts' => LicenseContract::query()
                ->whereNotIn('status', ['Terminated']),
        };
    }
}
