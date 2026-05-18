<?php

namespace App\Mail;

use App\Models\LicenseContract;
use Illuminate\Bus\Queueable;
use Illuminate\Mail\Mailable;
use Illuminate\Mail\Mailables\Content;
use Illuminate\Mail\Mailables\Envelope;
use Illuminate\Queue\SerializesModels;

class LicenseExpiringMail extends Mailable
{
    use Queueable, SerializesModels;

    public function __construct(
        public LicenseContract $license,
        public int $daysRemaining,
        public bool $isTest = false,
    ) {
    }

    public function envelope(): Envelope
    {
        $prefix = $this->isTest ? '[TEST] ' : '';

        return new Envelope(
            subject: "{$prefix}[ITAMS] License Expiry Reminder: {$this->license->software_name} expires in {$this->daysRemaining} day(s)",
        );
    }

    public function content(): Content
    {
        return new Content(
            view: 'emails.license-expiring',
            with: [
                'license'       => $this->license,
                'daysRemaining' => $this->daysRemaining,
                'isTest'        => $this->isTest,
            ],
        );
    }
}
