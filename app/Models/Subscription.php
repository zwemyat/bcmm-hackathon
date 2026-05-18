<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Subscription extends Model
{
    public const CURRENCIES = [
        'MMK' => 'Myanmar (MMK)',
        'JPY' => 'Japan (JPY)',
        'USD' => 'USD',
    ];

    /**
     * `reminder_date` deliberately omitted from fillable/casts. The column
     * still exists in the schema (`2026_05_13_000002_create_subscriptions_table`)
     * but is no longer maintained — the previous `booted()` hook that wrote it
     * read from `mail_settings.reminder_days_before`, a field the current
     * notification engine doesn't honor. The reminder engine
     * (ExpiryNotificationCounter / CheckExpirations) drives everything from
     * NotificationSetting + the live `expire_date` instead. See audit M1/M3.
     */
    protected $fillable = [
        'service_type', 'project_name', 'subscription_name', 'vendor_name', 'status',
        'period', 'previous_cost', 'expire_date', 'renewal_cost', 'currency',
        'renewal_type', 'renewal_status', 'remarks', 'modified_by',
    ];

    protected $casts = [
        'expire_date' => 'date',
        'previous_cost' => 'decimal:2',
        'renewal_cost' => 'decimal:2',
    ];
}
