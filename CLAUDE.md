# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

ITAMS — an IT Asset Management System for tracking PCs, network devices, software licenses & contracts, and recurring service subscriptions, with automated expiry reminders. Server-rendered Laravel 11 + Blade + Bootstrap 5 app intended to run on-premise via XAMPP/Apache on Windows.

The spec doc `Renewal_Reminding_System 2.md` is **partly stale** (mentions only Subscription + PC modules, a 2-role permission model, and the path `C:/xampp/htdocs/rrs-system`). The actual system has four modules (PC, Devices, Subscriptions, Licenses & Contracts), a per-module × per-action permission grid, and lives at `D:\xampp\htdocs\itams`. Trust the code, not the spec, when they disagree. The `README.md` is the default Laravel stub — ignore it.

The `docs/` and `presentation/` directories contain Python generators (`build_*.py` + their .docx / .xlsx / .pptx outputs) — these are **deliverables, not part of the application code**. Don't search them when grepping the codebase for app behaviour.

## Environment & commands

This is a Windows / XAMPP setup. `php` and `composer` are not on PATH; invoke them explicitly:

- PHP: `D:\xampp\php\php.exe`
- Composer: `composer.phar` at the project root (PowerShell: `& "D:\xampp\php\php.exe" composer.phar <cmd>`)
- App URL (Apache vhost via XAMPP `htdocs`): `http://localhost/itams/public`

No frontend build — Bootstrap 5.3, Bootstrap Icons, and (login page) Inter font are loaded from CDN in the Blade layouts. There is no `package.json`, no Vite, no `npm install`.

Common commands (PowerShell syntax):

```powershell
& "D:\xampp\php\php.exe" artisan migrate                    # apply DB migrations
& "D:\xampp\php\php.exe" artisan migrate:fresh --seed       # wipe + reseed
& "D:\xampp\php\php.exe" artisan app:check-expirations      # run the staggered digest send (idempotency: NONE — see below)
& "D:\xampp\php\php.exe" artisan schedule:work              # run the scheduler in the foreground (local testing)
& "D:\xampp\php\php.exe" artisan test                       # PHPUnit via the Laravel wrapper
& "D:\xampp\php\php.exe" artisan test --filter=TestName     # single test
& "D:\xampp\php\php.exe" artisan tinker                     # REPL
```

In production the daily scheduler is wired via Windows Task Scheduler running `php artisan schedule:run` every minute (see `routes/console.php`, which schedules `app:check-expirations` daily at 09:00).

### One-time setup: public storage link (required for avatars)

Avatars upload to `storage/app/public/avatars/...` and are served via `asset('storage/...')`, which expects `public/storage` to exist as a symlink/junction. **Without it, uploads succeed silently but the `<img>` 404s and the gradient initial keeps showing.** Set up once per checkout:

```powershell
# Preferred — needs Admin shell OR Windows Developer Mode:
& "D:\xampp\php\php.exe" artisan storage:link

# Fallback — no admin required, works the same:
cmd /c mklink /J "D:\xampp\htdocs\itams\public\storage" "D:\xampp\htdocs\itams\storage\app\public"
```

## Architecture you need before editing

### Permission model — per-module × per-action

Admin bypass + per-module/per-action boolean grid:

- `users.role === 'admin'` → `$user->isAdmin()` short-circuits every check to `true`.
- Non-admins: each of four modules has **separate** `can_view_<module>` and `can_edit_<module>` boolean columns. Modules declared in `User::MODULES`: `pc_assets`, `subscriptions`, `licenses_contracts`, `devices`. **Edit implies view** (`canView()` returns true if either bool is set).
- Use `$user->canAccess($module, 'view'|'edit')` in Blade/controllers.

Middleware aliases registered in `bootstrap/app.php`:

- `admin` — `EnsureUserIsAdmin` (admin-only routes: user management, mail settings, notification settings, activity logs).
- `module:<name>,<view|edit>` — `EnsureUserCanAccessModule`. The `routes/web.php` pattern for every module is two groups: a `view` group containing `index`/`show`/`export`/`template`, and an `edit` group containing write/import/bulk routes. Mirror this when adding modules.

### Notifications / expiry — live computation, not a notifications table

**The historical `notifications` table was dropped** (`2026_05_17_000002_drop_notifications_table.php`); a previous design that stored a Notification row per reminder is gone. The current system computes everything live.

Three pieces work together:

1. **`NotificationSetting` (DB table `notification_settings`, one row per module)** — per-module config:
   - `enabled` flag (must be true for that module's notifications to appear or be sent).
   - `days_before_set` (JSON array, allowed values `{10, 20, 30}`). `windowDays()` returns `max(set)` (drives the bell page and badge — "anything within N days"). `selectedDays()` returns the descending unique list (drives the staggered digest send).
   - `recipients` textarea, parsed by `recipientsArray()`. Empty → fall back to all admin user emails (see `CheckExpirations::sendStaggeredFor`).
   - `mail_settings.reminder_days_before` and `mail_settings.reminder_recipients` are **vestigial** — the form still saves them but `CheckExpirations` and `ExpiryNotificationCounter` read from `notification_settings`. Don't add new reminder behaviour to `mail_settings`.

2. **`ExpiryNotificationCounter::summary($user)`** (`app/Support/ExpiryNotificationCounter.php`) — the oracle for the bell badge and the notifications page KPIs. For each module that's `enabled`, it pulls live rows from `Subscription` / `LicenseContract` whose `expire_date <= today + windowDays()`, subtracts rows already read by `$user`, and buckets the remainder into `overdue` (`days < 0`) / `due_soon` (`0..7`) / `upcoming` (`>7`). PC Master and Devices are declared in `NotificationSetting::MODULES` but **only subscriptions and licenses_contracts** are actually scanned — adding more requires extending the counter and command.

3. **`NotificationRead` (DB table `notification_reads`)** — per-user read tracking, NOT per-notification (there are no notifications). Unique key: `(user_id, module, notifiable_id)`. Crucially, each row stores a `read_signature` = `NotificationRead::signature($expireDate, $daysRemaining)` = `"YYYY-MM-DD|<bucket>"` where bucket is `overdue` / `soon` / `upcoming`. A stored read is honoured **only while the live signature still matches** — if the underlying record's expire_date or urgency-bucket shifts (e.g., something previously marked read crosses from `upcoming` into `soon`), the item re-surfaces as unread. Preserve this signature scheme if you touch `markRead` or the counter.

### Expiration command — staggered digests

`app:check-expirations` (`app/Console/Commands/CheckExpirations.php`) runs daily at 09:00. Its model is **completely different from the old per-row reminder**:

1. **Mark past-due subscriptions Expired** (status flip on `renewal_status`). Licenses do NOT get auto-marked Expired here.
2. For each module × each day in `NotificationSetting->selectedDays()`: find rows where `expire_date == today + N days`, send **one `ExpiryReminderDigest` per bucket** (not per row) to that module's recipients. Subscriptions additionally get `renewal_status = Pending` flipped on rows included in a digest; licenses don't.
3. **No idempotency.** Re-running the command on the same day re-sends the same digests. There is no per-day dedupe table. If a send fails, retry; if it succeeded once, don't re-run unless you want a duplicate.

The mailable is `App\Mail\ExpiryReminderDigest` (single digest containing all records in one day-bucket). Older per-item mailables (`SubscriptionExpiringMail`, `LicenseExpiringMail`) still exist in `app/Mail/` but are not called from the command — they're effectively dead code unless re-wired. `App\Mail\UserCredentialsMail` is sent live when an admin creates a new user (see `UserController::store`); `App\Mail\PasswordResetMail` exists but is not wired to any route.

### Mail config has two sources

SMTP credentials come either from `.env` (default) OR from the `mail_settings` DB row.

- `MailSetting::current()` always returns the singleton row (firstOrCreate).
- When `mail_settings.enabled === true`, callers must push runtime overrides via `config([...])` and call `Mail::purge('smtp')` before sending — see `MailSettingController::sendTest` for the canonical sequence. **`CheckExpirations` does NOT do this purge**, so if you rely on DB-SMTP for scheduled reminders, you'd need to add the same sequence there (or use `.env`).
- `mail_settings.password` is cast `encrypted`.
- `MailSettingController::sendTest` sends a plain `Mail::raw` test message — not a preview of the digest format.

### Activity log

`App\Support\ActivityLogger::log($action, $description, $subject = null, $properties = [], $overrides = [])` is the single entry point. Controllers call it on login, logout, login_failed, created, updated, deleted, imported, renewed, mail_test, mail_sent, and similar. Captures `subject_type`/`subject_id` polymorphically, plus IP and user-agent. Failed logins use `$overrides` to record the attempted email without a user_id.

### Layout & theming conventions

- All authenticated pages `@extends('layouts.app')`. The layout renders sidebar + topbar **only when `@auth`**; guest views (login) yield content directly into `<body>`. The login view is wholly responsible for its own background and chrome.
- Theme persistence: `localStorage.rrs.theme` (`"light"` | `"dark"`). The inline script in `<head>` applies it before CSS loads — preserve this if you rewrite the layout, otherwise dark-mode users get a flash. The auth screen's theme toggle writes to the same key so the choice survives login.
- No central CSS file. Each view embeds its own `<style>` block. The `.glass-card`, `.kpi-card`, `.stat-row` / `.stat-cell`, `.btn-icon-soft`, `.quick-action`, `.status-chip`, `.live-pill`, `.module-card` patterns from `layouts.app` and the major index pages are reused everywhere — match them rather than inventing new ones.
- Clipboard-API uses inside Bootstrap dropdowns (e.g. the "Copy email" button in the user menu) need a `document.execCommand('copy')` fallback and must capture DOM refs **before** the `await navigator.clipboard.writeText(...)` call — otherwise focus loss when the dropdown closes makes the Promise reject silently.

### Side-effect on Subscription save

`Subscription::booted()` auto-recomputes `reminder_date = expire_date - reminder_days_before` on every save, reading the window from `MailSetting` (still — this one path predates the per-module settings split). The hook silently swallows errors so initial migrations don't break.

## Things that have bitten in the past

- **The `public/storage` link is not in git.** A fresh checkout has avatar uploads that succeed server-side but never display. See "One-time setup" above.
- **Two migrations share the timestamp `2026_05_17_000001`** (`split_module_permissions_into_view_and_edit` and `add_license_contract_id_to_notifications_table`). Run order is alphabetical within a timestamp, but the second was promptly nullified by `2026_05_17_000002_drop_notifications_table`. Don't assume timestamps are unique; don't try to depend on the dropped `notifications` table.
- **`mail_settings.reminder_days_before` and `mail_settings.reminder_recipients`** look authoritative but are unused by the current notification flow. The form still validates them; expect to either delete them or migrate the values to `notification_settings`.
- Calling the helper `composer` directly fails — use `& "D:\xampp\php\php.exe" composer.phar ...`.
- Don't redirect native-exe stderr inside PowerShell with `2>&1` — it wraps each line in an ErrorRecord and sets `$?` to false even on exit 0.
- When sending mail via the DB-SMTP path, forget the `Mail::purge('smtp')` call and you'll send with stale config from a previous request.
- `app:check-expirations` is **not idempotent**. Re-running on the same day double-sends every digest that has recipients.
