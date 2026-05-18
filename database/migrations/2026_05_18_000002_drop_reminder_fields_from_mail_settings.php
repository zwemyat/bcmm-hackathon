<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Drops `reminder_recipients` and `reminder_days_before` from
     * `mail_settings`. These were superseded by the per-module
     * `notification_settings` table (recipients + days_before_set) but the old
     * columns lingered as vestigial state that the form still let admins edit.
     * See audit M3.
     */
    public function up(): void
    {
        Schema::table('mail_settings', function (Blueprint $table) {
            if (Schema::hasColumn('mail_settings', 'reminder_recipients')) {
                $table->dropColumn('reminder_recipients');
            }
            if (Schema::hasColumn('mail_settings', 'reminder_days_before')) {
                $table->dropColumn('reminder_days_before');
            }
        });
    }

    public function down(): void
    {
        Schema::table('mail_settings', function (Blueprint $table) {
            $table->text('reminder_recipients')->nullable()->after('from_name');
            $table->unsignedInteger('reminder_days_before')->default(30)->after('reminder_recipients');
        });
    }
};
