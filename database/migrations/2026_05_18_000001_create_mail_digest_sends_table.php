<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::create('mail_digest_sends', function (Blueprint $table) {
            $table->id();
            $table->string('module', 64);
            // Day-mark this digest fired for (e.g. 30, 20, 10). Comes from
            // NotificationSetting::selectedDays().
            $table->unsignedSmallInteger('day_mark');
            $table->date('sent_on');
            $table->unsignedInteger('records_count')->default(0);
            $table->unsignedInteger('recipients_count')->default(0);
            $table->timestamp('sent_at');
            $table->timestamps();

            // Idempotency key — app:check-expirations refuses to re-send a
            // (module, day_mark) pair on the same day unless --force is passed.
            $table->unique(['module', 'day_mark', 'sent_on']);
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('mail_digest_sends');
    }
};
