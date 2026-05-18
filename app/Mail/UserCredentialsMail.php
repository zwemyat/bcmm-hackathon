<?php

namespace App\Mail;

use App\Models\User;
use Illuminate\Bus\Queueable;
use Illuminate\Mail\Mailable;
use Illuminate\Mail\Mailables\Content;
use Illuminate\Mail\Mailables\Envelope;
use Illuminate\Queue\SerializesModels;

/**
 * Welcome email sent when an admin creates a new user. Carries a single-use
 * password-broker token URL — NOT a cleartext password. The user clicks the
 * link to set their own password through the standard reset UI.
 */
class UserCredentialsMail extends Mailable
{
    use Queueable, SerializesModels;

    public function __construct(
        public User $user,
        public string $setupUrl,
        public int $expireMinutes,
        public string $loginUrl,
    ) {
    }

    public function envelope(): Envelope
    {
        $app = config('app.name', 'ITAMS');
        return new Envelope(
            subject: "[{$app}] Welcome — set your password to get started",
        );
    }

    public function content(): Content
    {
        return new Content(
            view: 'emails.user-credentials',
            with: [
                'user'          => $this->user,
                'setupUrl'      => $this->setupUrl,
                'expireMinutes' => $this->expireMinutes,
                'loginUrl'      => $this->loginUrl,
                'appName'       => config('app.name', 'ITAMS'),
            ],
        );
    }
}
