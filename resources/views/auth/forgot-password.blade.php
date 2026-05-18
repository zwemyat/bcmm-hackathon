@extends('layouts.app')

@section('title', 'Forgot password')

@section('content')
@php
    // C3: a single 'submitted' state for any successful POST, regardless of
    // whether the email belongs to an admin. The page renders the same content
    // either way so it can't be probed for account-existence / role.
    $submittedEmail = session('submitted_email');
    $state = $submittedEmail ? 'submitted' : 'form';
@endphp

<div class="fp-bg">
    {{-- Decorative background blobs --}}
    <div class="fp-blob fp-blob--a"></div>
    <div class="fp-blob fp-blob--b"></div>
    <div class="fp-blob fp-blob--c"></div>

    <div class="fp-shell">
        <div class="fp-card fp-card--{{ $state }}">

            @if($state === 'submitted')
                {{-- ============ SUBMITTED (admin or not — same UI) ============ --}}
                <div class="fp-hero">
                    <div class="fp-hero-icon fp-hero-icon--success">
                        <i class="bi bi-envelope-check-fill"></i>
                    </div>
                </div>
                <h1 class="fp-title">Request received</h1>
                <p class="fp-sub">
                    If an admin account is registered with this email, we've sent a secure reset link.
                </p>
                <div class="fp-email-chip">
                    <i class="bi bi-envelope"></i>
                    <span>{{ $submittedEmail }}</span>
                </div>
                <p class="fp-fineprint">
                    Don't see it within a few minutes? Standard accounts can't reset by email &mdash;
                    please contact your administrator.
                </p>

                @if(!empty($adminEmails))
                    @php
                        $subject = rawurlencode('Password reset request');
                        $body    = rawurlencode("Hi,\n\nI need help resetting my password for " . config('app.name', 'ITAMS') . ".\nMy email: " . $submittedEmail . "\n\nThank you.");
                        $primary = $adminEmails[0];
                        $cc      = count($adminEmails) > 1 ? implode(',', array_slice($adminEmails, 1)) : null;
                        $mailto  = "mailto:{$primary}?" . ($cc ? "cc=" . rawurlencode($cc) . "&" : '') . "subject={$subject}&body={$body}";
                    @endphp
                    <a href="{{ $mailto }}" class="fp-btn fp-btn--primary">
                        <i class="bi bi-envelope"></i> Email administrator
                    </a>
                @endif

                <a href="{{ route('login') }}" class="fp-link-back">
                    <i class="bi bi-arrow-left"></i> Back to sign in
                </a>

            @else
                {{-- ============ FORM ============ --}}
                <div class="fp-hero">
                    <div class="fp-hero-icon">
                        <i class="bi bi-key-fill"></i>
                        <span class="fp-hero-ring"></span>
                    </div>
                </div>
                <h1 class="fp-title">Forgot password</h1>
                <p class="fp-sub">We'll email you a secure reset link.</p>

                @if($errors->any())
                    <div class="fp-alert fp-alert--error">
                        <i class="bi bi-exclamation-triangle-fill"></i>
                        <span>{{ $errors->first() }}</span>
                    </div>
                @endif

                <form method="POST" action="{{ route('password.email') }}" id="forgotForm" class="fp-form">
                    @csrf
                    <label class="fp-float">
                        <input type="email" id="email" name="email"
                               value="{{ old('email') }}"
                               placeholder=" "
                               required autofocus autocomplete="email">
                        <span class="fp-float-label">Email address</span>
                        <i class="bi bi-envelope fp-float-icon"></i>
                    </label>

                    <button type="submit" class="fp-btn fp-btn--primary fp-submit">
                        <span class="default">Send reset link</span>
                        <span class="loading d-none"><span class="spinner-border spinner-border-sm me-1"></span> Sending…</span>
                    </button>
                </form>

                <div class="fp-divider"><span>or</span></div>

                <p class="fp-helper">
                    Standard users &mdash; please <a href="{{ route('password.request') }}#contact" id="askAdminLink">contact your administrator</a>.
                </p>

                <a href="{{ route('login') }}" class="fp-link-back">
                    <i class="bi bi-arrow-left"></i> Back to sign in
                </a>
            @endif
        </div>

        <div class="fp-foot">{{ config('app.name', 'ITAMS') }} &middot; &copy; {{ date('Y') }}</div>
    </div>
</div>

@include('auth._fp-shell')

<script>
    (function () {
        const form   = document.getElementById('forgotForm');
        const submit = form?.querySelector('.fp-submit');
        if (form && submit) {
            form.addEventListener('submit', () => {
                submit.disabled = true;
                submit.querySelector('.default')?.classList.add('d-none');
                submit.querySelector('.loading')?.classList.remove('d-none');
            });
        }

        // "Contact your administrator" inline link — submit empty form to trigger contact-admin path
        const askLink = document.getElementById('askAdminLink');
        if (askLink) {
            askLink.addEventListener('click', (e) => {
                e.preventDefault();
                const emailInput = document.getElementById('email');
                if (emailInput && emailInput.value.trim()) {
                    form?.submit();
                } else {
                    emailInput?.focus();
                    emailInput?.setAttribute('placeholder', 'Enter your email first');
                }
            });
        }
    })();
</script>
@endsection
