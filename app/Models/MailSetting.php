<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\Mail;

class MailSetting extends Model
{
    protected $fillable = [
        'mailer', 'host', 'port', 'encryption', 'auth_mode', 'username',
        'password', 'from_address', 'from_name', 'enabled',
        'reminder_recipients', 'reminder_days_before',
    ];

    protected $casts = [
        'password' => 'encrypted',
        'enabled' => 'boolean',
        'port' => 'integer',
        'reminder_days_before' => 'integer',
    ];

    public static function current(): self
    {
        return static::firstOrCreate([], [
            'mailer' => 'smtp',
            'port' => 587,
            'enabled' => false,
            'reminder_days_before' => 30,
        ]);
    }

    public function recipientsArray(): array
    {
        if (! $this->reminder_recipients) {
            return [];
        }

        return collect(preg_split('/[\s,;]+/', $this->reminder_recipients))
            ->map(fn ($e) => trim($e))
            ->filter(fn ($e) => $e !== '' && filter_var($e, FILTER_VALIDATE_EMAIL))
            ->values()
            ->toArray();
    }

    /**
     * If DB-stored SMTP is the active source (enabled = true), push these
     * settings onto the runtime config and purge Mail's transport cache so the
     * next send() picks them up. No-op when disabled — caller keeps using the
     * .env-driven default mailer.
     *
     * Idempotent: safe to call once per request / once per scheduled run.
     */
    public function applyRuntimeConfig(): void
    {
        if (! $this->enabled) {
            return;
        }

        $mailer = $this->mailer ?: 'smtp';

        config([
            'mail.default'                 => $mailer,
            'mail.mailers.smtp.host'       => $this->host,
            'mail.mailers.smtp.port'       => $this->port,
            'mail.mailers.smtp.encryption' => $this->encryption,
            'mail.mailers.smtp.auth_mode'  => $this->auth_mode,
            'mail.mailers.smtp.username'   => $this->username,
            'mail.mailers.smtp.password'   => $this->password,
            'mail.from.address'            => $this->from_address ?: config('mail.from.address'),
            'mail.from.name'               => $this->from_name ?: config('mail.from.name'),
        ]);

        Mail::purge($mailer);
        if ($mailer !== 'smtp') {
            Mail::purge('smtp');
        }
    }
}
