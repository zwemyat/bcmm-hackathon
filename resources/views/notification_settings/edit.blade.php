@extends('layouts.app')
@section('title', 'Notification Settings')
@section('content')
@php
    $modules = \App\Models\NotificationSetting::MODULES;
    $moduleIcons = [
        'subscriptions'      => 'bi-calendar-event',
        'licenses_contracts' => 'bi-file-earmark-text',
    ];
    $moduleHints = [
        'subscriptions'      => 'Renewal reminders based on subscription expire date.',
        'licenses_contracts' => 'Renewal reminders based on license / contract expire date.',
    ];
    $activeTab = session('active_tab', 'subscriptions');
    if (! array_key_exists($activeTab, $modules)) $activeTab = 'subscriptions';
@endphp

<div class="page-header">
    <div>
        <h1 class="page-title">Notification Settings</h1>
        <div class="page-subtitle">Configure per-module reminder emails &mdash; window, recipients, and on/off.</div>
    </div>
    <a href="{{ route('mail-settings.edit') }}" class="quick-action">
        <i class="bi bi-envelope-fill"></i> Mail Settings
    </a>
</div>

<div class="card">
    <div class="card-body p-0">
        <ul class="nav nav-tabs notification-tabs" role="tablist">
            @foreach($modules as $key => $label)
                @php
                    $setting = $settings[$key] ?? null;
                    $isActive = $key === $activeTab;
                @endphp
                <li class="nav-item" role="presentation">
                    <button class="nav-link {{ $isActive ? 'active' : '' }}" id="tab-{{ $key }}"
                            data-bs-toggle="tab" data-bs-target="#pane-{{ $key }}" type="button"
                            role="tab" aria-controls="pane-{{ $key }}" aria-selected="{{ $isActive ? 'true' : 'false' }}">
                        <i class="bi {{ $moduleIcons[$key] }}"></i>
                        <span class="ms-1">{{ $label }}</span>
                        @if($setting && $setting->enabled)
                            <span class="badge bg-success-subtle text-success-emphasis ms-1" title="Notifications enabled">on</span>
                        @endif
                    </button>
                </li>
            @endforeach
        </ul>

        <div class="tab-content p-4">
            @foreach($modules as $key => $label)
                @php
                    $setting = $settings[$key] ?? null;
                    $isActive = $key === $activeTab;
                    $errorKey = "recipients_{$key}";
                @endphp
                <div class="tab-pane fade {{ $isActive ? 'show active' : '' }}" id="pane-{{ $key }}" role="tabpanel" aria-labelledby="tab-{{ $key }}">

                        <div class="d-flex align-items-center gap-3 mb-3">
                            <span class="module-tab-icon"><i class="bi {{ $moduleIcons[$key] }}"></i></span>
                            <div>
                                <div class="fw-semibold">{{ $label }}</div>
                                <div class="text-muted small">{{ $moduleHints[$key] }}</div>
                            </div>
                        </div>

                        <form method="POST" action="{{ route('notification-settings.update', $key) }}">
                            @csrf @method('PUT')

                            {{-- Enable / source toggle (reuses mail-enable-card pattern) --}}
                            <div class="mail-enable-card card mb-3 {{ $setting->enabled ? 'is-on' : '' }}">
                                <div class="card-body d-flex align-items-center gap-3">
                                    <div class="mail-enable-icon">
                                        <i class="bi bi-bell"></i>
                                    </div>
                                    <div class="flex-grow-1">
                                        <div class="fw-semibold">Send reminders for {{ $label }}</div>
                                        <div class="text-muted small">When off, no notifications or emails are generated for this module.</div>
                                    </div>
                                    <div class="form-check form-switch m-0">
                                        <input type="hidden" name="enabled" value="0">
                                        <input type="checkbox" name="enabled" value="1" id="enabled_{{ $key }}" class="form-check-input notification-enable-toggle" role="switch" @checked($setting->enabled) style="width: 3rem; height: 1.6rem;">
                                    </div>
                                </div>
                            </div>

                            <div class="row g-3">
                                <div class="col-md-4">
                                    <label class="form-label">Reminder days before expiry <span class="text-danger">*</span></label>
                                    @php
                                        $presets = [30, 20, 10];
                                        // Pre-check items already in the saved set; fall back to nearest
                                        // preset if existing data is e.g. legacy 15.
                                        $rawOld   = old('days_before_set', $setting->days_before_set ?? []);
                                        $rawOld   = is_array($rawOld) ? $rawOld : [];
                                        $currentSet = array_values(array_intersect($presets, array_map('intval', $rawOld)));
                                        if (empty($currentSet) && ! empty($rawOld)) {
                                            // Legacy single value (e.g. 15) — pick the nearest preset, tiebreak larger.
                                            $legacy = (int) $rawOld[0];
                                            $currentSet = [collect($presets)
                                                ->sortBy(fn ($v) => sprintf('%05d-%05d', abs($v - $legacy), 9999 - $v))
                                                ->first()];
                                        } elseif (empty($currentSet)) {
                                            $currentSet = [30];
                                        }
                                    @endphp
                                    <div class="day-mark-list">
                                        @foreach($presets as $d)
                                            @php $on = in_array($d, $currentSet, true); @endphp
                                            <label class="day-mark-card {{ $on ? 'is-on' : '' }}" for="days_{{ $key }}_{{ $d }}">
                                                <span class="day-mark-meta">
                                                    <span class="day-mark-value">{{ $d }}</span>
                                                    <span class="day-mark-unit">days before</span>
                                                </span>
                                                <span class="form-check form-switch m-0">
                                                    <input type="checkbox"
                                                           name="days_before_set[]"
                                                           id="days_{{ $key }}_{{ $d }}"
                                                           value="{{ $d }}"
                                                           class="form-check-input day-mark-toggle"
                                                           role="switch"
                                                           @checked($on)>
                                                </span>
                                            </label>
                                        @endforeach
                                    </div>
                                    <small class="text-muted d-block mt-2">Toggle one or more. The widest enabled window drives the badge &amp; list.</small>
                                    @error('days_before_set')<div class="invalid-feedback d-block">{{ $message }}</div>@enderror
                                    @error('days_before_set.*')<div class="invalid-feedback d-block">{{ $message }}</div>@enderror
                                </div>
                                <div class="col-md-8">
                                    <label class="form-label">Recipients</label>
                                    <textarea name="recipients" rows="3" class="form-control @error($errorKey) is-invalid @enderror" placeholder="One or more emails — separate with comma, semicolon, or newline.&#10;e.g. ops@company.com, admin@company.com">{{ old('recipients', $setting->recipients) }}</textarea>
                                    <small class="text-muted">Leave empty to fall back to all admin users' emails.</small>
                                    @error($errorKey)<div class="invalid-feedback d-block">{{ $message }}</div>@enderror
                                </div>
                            </div>

                            <div class="d-flex gap-2 mt-3">
                                <button class="btn btn-primary"><i class="bi bi-check2"></i> Save {{ $label }} settings</button>
                            </div>
                        </form>
                </div>
            @endforeach
        </div>
    </div>
</div>

<div class="text-muted small mt-3">
    <i class="bi bi-info-circle"></i>
    Notifications are sent by the daily <code>app:check-expirations</code> command at 09:00. Email transport is configured in
    <a href="{{ route('mail-settings.edit') }}">Mail Settings</a>.
</div>

<style>
    .notification-tabs {
        border-bottom: 1px solid rgba(31, 38, 135, 0.08);
        padding: .5rem .5rem 0;
        margin-bottom: 0;
        gap: .25rem;
    }
    .notification-tabs .nav-link {
        border: 1px solid transparent;
        border-radius: .55rem .55rem 0 0;
        color: #475569;
        font-size: .88rem;
        font-weight: 500;
        padding: .55rem .85rem;
        display: inline-flex;
        align-items: center;
        gap: .25rem;
    }
    .notification-tabs .nav-link:hover {
        color: #0d6efd;
        background: rgba(13, 110, 253, 0.04);
        border-color: transparent;
    }
    .notification-tabs .nav-link.active {
        color: #0d6efd;
        background: #fff;
        border-color: rgba(31, 38, 135, 0.08) rgba(31, 38, 135, 0.08) #fff;
        font-weight: 600;
    }

    .module-tab-icon {
        width: 40px; height: 40px;
        border-radius: .55rem;
        background: rgba(13, 110, 253, 0.1);
        color: #0d6efd;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 1.2rem;
    }

    /* Day-mark on/off cards: one per preset (30/20/10 days). */
    .day-mark-list {
        display: flex;
        flex-direction: column;
        gap: .4rem;
    }
    .day-mark-card {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: .75rem;
        padding: .55rem .85rem;
        border-radius: .6rem;
        border: 1px solid rgba(31, 38, 135, 0.12);
        background: #fff;
        cursor: pointer;
        margin: 0;
        transition: background .15s ease, border-color .15s ease, box-shadow .15s ease;
    }
    .day-mark-card:hover {
        border-color: rgba(13, 110, 253, 0.3);
        background: rgba(13, 110, 253, 0.025);
    }
    .day-mark-card.is-on {
        border-color: rgba(13, 110, 253, 0.45);
        background: rgba(13, 110, 253, 0.06);
        box-shadow: 0 1px 2px rgba(13, 110, 253, 0.08);
    }
    .day-mark-meta { display: inline-flex; align-items: baseline; gap: .4rem; line-height: 1.1; }
    .day-mark-value { font-size: 1.05rem; font-weight: 700; color: #0f172a; }
    .day-mark-unit  { font-size: .72rem; color: #64748b; font-weight: 500; }
    .day-mark-card.is-on .day-mark-value { color: #0d6efd; }
    .day-mark-card .form-check-input {
        width: 2.2rem; height: 1.25rem;
        margin: 0;
        cursor: pointer;
    }

    [data-bs-theme="dark"] .day-mark-card {
        background: rgba(255, 255, 255, 0.04);
        border-color: rgba(255, 255, 255, 0.08);
    }
    [data-bs-theme="dark"] .day-mark-card:hover {
        background: rgba(147, 197, 253, 0.05);
        border-color: rgba(147, 197, 253, 0.3);
    }
    [data-bs-theme="dark"] .day-mark-card.is-on {
        background: rgba(147, 197, 253, 0.1);
        border-color: rgba(147, 197, 253, 0.45);
        box-shadow: 0 1px 2px rgba(147, 197, 253, 0.1);
    }
    [data-bs-theme="dark"] .day-mark-value { color: #f1f5f9; }
    [data-bs-theme="dark"] .day-mark-card.is-on .day-mark-value { color: #93c5fd; }
    [data-bs-theme="dark"] .day-mark-unit { color: #94a3b8; }

    [data-bs-theme="dark"] .notification-tabs { border-bottom-color: rgba(255, 255, 255, 0.06); }
    [data-bs-theme="dark"] .notification-tabs .nav-link { color: #cfd8dc; }
    [data-bs-theme="dark"] .notification-tabs .nav-link:hover { background: rgba(147, 197, 253, 0.06); color: #93c5fd; }
    [data-bs-theme="dark"] .notification-tabs .nav-link.active {
        background: #1a1f29;
        color: #93c5fd;
        border-color: rgba(255, 255, 255, 0.06) rgba(255, 255, 255, 0.06) #1a1f29;
    }
    [data-bs-theme="dark"] .module-tab-icon { background: rgba(147, 197, 253, 0.15); color: #93c5fd; }
</style>

<script>
    (function () {
        // Module-level enable card
        document.querySelectorAll('.notification-enable-toggle').forEach(input => {
            const card = input.closest('.mail-enable-card');
            input.addEventListener('change', () => {
                card?.classList.toggle('is-on', input.checked);
            });
        });

        // Per-day-mark on/off switches
        document.querySelectorAll('.day-mark-toggle').forEach(input => {
            const card = input.closest('.day-mark-card');
            input.addEventListener('change', () => {
                card?.classList.toggle('is-on', input.checked);
            });
        });
    })();
</script>
@endsection
