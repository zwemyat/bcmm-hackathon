<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

/**
 * Idempotency ledger for app:check-expirations. One row per successfully-sent
 * (module, day_mark, date) digest. CheckExpirations skips a bucket whose row
 * already exists for today unless --force is passed.
 */
class MailDigestSend extends Model
{
    protected $fillable = [
        'module', 'day_mark', 'sent_on', 'records_count', 'recipients_count', 'sent_at',
    ];

    protected $casts = [
        'sent_on'          => 'date',
        'sent_at'          => 'datetime',
        'day_mark'         => 'integer',
        'records_count'    => 'integer',
        'recipients_count' => 'integer',
    ];
}
