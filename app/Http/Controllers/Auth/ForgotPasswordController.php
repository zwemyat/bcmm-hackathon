<?php

namespace App\Http\Controllers\Auth;

use App\Http\Controllers\Controller;
use App\Mail\PasswordResetMail;
use App\Models\User;
use App\Support\ActivityLogger;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Hash;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Mail;
use Illuminate\Support\Facades\Password;
use Illuminate\Support\Str;
use Illuminate\Validation\Rules\Password as PasswordRule;

class ForgotPasswordController extends Controller
{
    /**
     * Generic error shown on any failed reset attempt. Same wording for invalid
     * token, expired token, non-admin user, or unknown email — so the page can't
     * be used to enumerate which addresses are admin accounts.
     */
    private const GENERIC_RESET_ERROR = 'This reset link is invalid or has expired. Please request a new one.';

    public function showLinkRequestForm()
    {
        return view('auth.forgot-password', [
            'adminEmails' => $this->adminEmails(),
        ]);
    }

    /**
     * Returns admin email addresses for the "contact your administrator"
     * mailto link on the forgot-password page. Internal tool — exposing the
     * IT admin's address is acceptable; users could already find them in any
     * sent email from the app.
     */
    private function adminEmails(): array
    {
        return User::query()
            ->where('role', 'admin')
            ->pluck('email')
            ->filter()
            ->values()
            ->all();
    }

    public function sendResetLinkEmail(Request $request)
    {
        $data = $request->validate([
            'email' => ['required', 'email'],
        ]);

        $user = User::query()->where('email', $data['email'])->first();

        // Two server-side paths (admin sends, non-admin denies) — but the
        // user-facing flash is identical so the page can't be probed for
        // "is this email an admin account?" enumeration. Differentiation
        // stays in the activity log for ops review.
        if (! $user || ! $user->isAdmin()) {
            ActivityLogger::log(
                action: 'password_reset_denied',
                description: "Password reset denied (non-admin or unknown) for {$data['email']}",
                overrides: [
                    'user_id' => null,
                    'user_name' => null,
                    'user_email' => $data['email'],
                ],
            );

            return back()->with('submitted_email', $data['email']);
        }

        $token = Password::broker()->createToken($user);
        $expireMinutes = (int) config('auth.passwords.users.expire', 60);
        $resetUrl = route('password.reset', ['token' => $token, 'email' => $user->email]);

        try {
            Mail::to($user->email)->send(new PasswordResetMail(
                user: $user,
                resetUrl: $resetUrl,
                expireMinutes: $expireMinutes,
            ));

            ActivityLogger::log(
                action: 'password_reset_requested',
                description: "Sent password reset link to admin {$user->email}",
                subject: $user,
            );
        } catch (\Throwable $e) {
            // Log the failure server-side, but do NOT surface "couldn't send to
            // <email>" to the user — that confirms the email is an admin. Show
            // the same generic submitted-state page either way; admins watching
            // the log will catch send failures.
            Log::warning('Failed to send password reset email', [
                'user_id' => $user->id,
                'email'   => $user->email,
                'error'   => $e->getMessage(),
            ]);
        }

        return back()->with('submitted_email', $data['email']);
    }

    public function showResetForm(Request $request, string $token)
    {
        return view('auth.reset-password', [
            'token' => $token,
            'email' => $request->query('email', ''),
        ]);
    }

    public function reset(Request $request)
    {
        $data = $request->validate([
            'token'    => ['required'],
            'email'    => ['required', 'email'],
            'password' => ['required', 'confirmed', PasswordRule::min(8)->mixedCase()->numbers()],
        ]);

        // H1: this endpoint is used for two paths — admin self-reset (token
        // came from /forgot-password) AND first-time setup for any newly-
        // created user (token came from UserController::store). Both need to
        // pass through. We don't gate by role here because the password broker
        // already verifies the token belongs to the user, and tokens can only
        // be minted by admin-authenticated code paths (the trigger endpoint
        // is admin-only; user creation is admin-only).
        $user = User::query()->where('email', $data['email'])->first();
        if (! $user) {
            return back()
                ->withInput($request->only('email'))
                ->withErrors(['email' => self::GENERIC_RESET_ERROR]);
        }

        $status = Password::broker()->reset(
            $data,
            function (User $user, string $password) {
                $user->forceFill([
                    'password'       => Hash::make($password),
                    'remember_token' => Str::random(60),
                ])->save();
            }
        );

        if ($status === Password::PASSWORD_RESET) {
            $isFirstTime = ! $user->isAdmin();
            ActivityLogger::log(
                action: $isFirstTime ? 'password_set' : 'password_reset',
                description: $isFirstTime
                    ? "User {$user->email} set their initial password via setup link"
                    : "Admin {$user->email} reset their password via email link",
                subject: $user,
            );

            return redirect()->route('login')->with('success', 'Password set. You can sign in with your new password.');
        }

        // Same generic error for invalid token / expired token / mismatched
        // user, so the page can't be used to enumerate token validity.
        return back()
            ->withInput($request->only('email'))
            ->withErrors(['email' => self::GENERIC_RESET_ERROR]);
    }
}
