<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        // This migration is a no-op on any modern checkout: the very next
        // migration (2026_05_17_000002_drop_notifications_table) drops the
        // entire notifications table, so adding a column here is wasted work.
        // Guarding on Schema::hasTable lets fresh installs and inconsistent
        // partial states (where the column was added but the migration row
        // wasn't recorded) both run cleanly. See audit M5 — this file is a
        // candidate for outright deletion once everyone has migrated past it.
        if (! Schema::hasTable('notifications') || Schema::hasColumn('notifications', 'license_contract_id')) {
            return;
        }

        Schema::table('notifications', function (Blueprint $table) {
            $table->dropForeign(['subscription_id']);
        });

        Schema::table('notifications', function (Blueprint $table) {
            $table->foreignId('subscription_id')->nullable()->change();
            $table->foreign('subscription_id')->references('id')->on('subscriptions')->cascadeOnDelete();

            $table->foreignId('license_contract_id')
                ->nullable()
                ->after('subscription_id')
                ->constrained('licenses_contracts')
                ->cascadeOnDelete();
        });
    }

    public function down(): void
    {
        if (! Schema::hasTable('notifications')) {
            return;
        }

        Schema::table('notifications', function (Blueprint $table) {
            $table->dropForeign(['license_contract_id']);
            $table->dropColumn('license_contract_id');
        });

        Schema::table('notifications', function (Blueprint $table) {
            $table->dropForeign(['subscription_id']);
            $table->foreignId('subscription_id')->nullable(false)->change();
            $table->foreign('subscription_id')->references('id')->on('subscriptions')->cascadeOnDelete();
        });
    }
};
