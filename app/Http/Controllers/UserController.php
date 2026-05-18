<?php

namespace App\Http\Controllers;

use App\Mail\UserCredentialsMail;
use App\Models\User;
use App\Support\ActivityLogger;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Mail;
use Illuminate\Support\Facades\Password as PasswordBroker;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Str;
use Illuminate\Validation\Rules\Password;

class UserController extends Controller
{
    private const PERM_FIELDS = [
        'can_view_pc_assets',
        'can_edit_pc_assets',
        'can_view_subscriptions',
        'can_edit_subscriptions',
        'can_view_licenses_contracts',
        'can_edit_licenses_contracts',
        'can_view_devices',
        'can_edit_devices',
    ];

    public function index(Request $request)
    {
        $query = User::query();

        if ($search = $request->get('search')) {
            $query->where(function ($q) use ($search) {
                $q->where('name', 'like', "%{$search}%")
                  ->orWhere('email', 'like', "%{$search}%");
            });
        }

        if ($role = $request->get('role')) {
            $query->where('role', $role);
        }

        $users = $query->orderBy('name')->paginate(20)->withQueryString();

        $kpis = [
            'total'  => User::count(),
            'admins' => User::where('role', 'admin')->count(),
            'users'  => User::where('role', 'user')->count(),
        ];

        return view('users.index', compact('users', 'kpis'));
    }

    public function create()
    {
        return view('users.create');
    }

    public function store(Request $request)
    {
        // H1: admin no longer enters a password. We auto-generate one that's
        // never communicated — it lives only as a bcrypt hash in the DB — and
        // email the new user a one-time setup link instead. The user picks
        // their actual password through the reset flow.
        $rules = [
            'name' => 'required|string|max:255',
            'email' => 'required|email|unique:users,email',
            'role' => 'required|in:admin,user',
            'avatar' => 'nullable|image|mimes:jpg,jpeg,png,webp|max:2048',
        ];
        foreach (self::PERM_FIELDS as $field) {
            $rules[$field] = 'sometimes|boolean';
        }

        $data = $request->validate($rules);

        if ($request->hasFile('avatar')) {
            $data['avatar'] = $request->file('avatar')->store('avatars', 'public');
        }

        $this->applyPermissions($request, $data);

        // Random throwaway. The model's 'password' => 'hashed' cast bcrypts it
        // on save, so nobody (not even the admin who created the account) ever
        // sees the cleartext.
        $data['password'] = Str::random(40);

        $user = User::create($data);

        // Mint a single-use token via Laravel's password broker. The setup
        // page is the same UI as the reset page — first-time setup is just a
        // reset where the placeholder password is unknown.
        $token         = PasswordBroker::broker()->createToken($user);
        $expireMinutes = (int) config('auth.passwords.users.expire', 60);
        $setupUrl      = route('password.reset', ['token' => $token, 'email' => $user->email]);

        ActivityLogger::log(
            action: 'created',
            description: "Created user {$user->name} ({$user->email}) with role {$user->role}",
            subject: $user,
        );

        $emailStatus = $this->sendCredentialsEmail($user, $setupUrl, $expireMinutes);

        $flash = $emailStatus === true
            ? "User created. Setup link sent to {$user->email}."
            : 'User created, but the setup email could not be sent: ' . $emailStatus;

        return redirect()->route('users.index')->with($emailStatus === true ? 'success' : 'warning', $flash);
    }

    public function edit(User $user)
    {
        return view('users.edit', compact('user'));
    }

    public function update(Request $request, User $user)
    {
        $rules = [
            'name' => 'required|string|max:255',
            'email' => "required|email|unique:users,email,{$user->id}",
            'password' => ['nullable', Password::min(8)->mixedCase()->numbers()],
            'role' => 'required|in:admin,user',
            'avatar' => 'nullable|image|mimes:jpg,jpeg,png,webp|max:2048',
        ];
        foreach (self::PERM_FIELDS as $field) {
            $rules[$field] = 'sometimes|boolean';
        }

        $data = $request->validate($rules);

        $this->applyPermissions($request, $data);

        if (empty($data['password'])) {
            unset($data['password']);
        }

        if ($request->hasFile('avatar')) {
            if ($user->avatar) {
                Storage::disk('public')->delete($user->avatar);
            }
            $data['avatar'] = $request->file('avatar')->store('avatars', 'public');
        } else {
            unset($data['avatar']);
        }

        $original = $user->only(array_keys($data));
        $user->update($data);

        $changes = collect($data)
            ->reject(fn ($v, $k) => ($original[$k] ?? null) == $v)
            ->reject(fn ($v, $k) => $k === 'password')
            ->keys()
            ->all();

        ActivityLogger::log(
            action: 'updated',
            description: "Updated user {$user->name} ({$user->email})",
            subject: $user,
            properties: ['changed_fields' => $changes, 'password_changed' => isset($data['password'])],
        );

        return redirect()->route('users.index')->with('success', 'User updated.');
    }

    public function destroy(User $user)
    {
        if ($user->id === auth()->id()) {
            return back()->with('error', 'You cannot delete your own account.');
        }

        // H3: refuse to delete the last admin — otherwise admin-only pages
        // (user management, mail settings, activity logs) become unreachable
        // and the system locks itself out.
        if ($user->isAdmin() && User::where('role', 'admin')->count() <= 1) {
            return back()->with('error', 'Cannot delete the last admin account.');
        }

        if ($user->avatar) {
            Storage::disk('public')->delete($user->avatar);
        }

        $label = "{$user->name} ({$user->email})";

        ActivityLogger::log(
            action: 'deleted',
            description: "Deleted user {$label}",
            subject: $user,
        );

        $user->delete();

        return redirect()->route('users.index')->with('success', 'User deleted.');
    }

    /**
     * Email the new user a one-time setup link instead of a cleartext
     * password. Returns true on success, or the error message string on
     * failure.
     */
    private function sendCredentialsEmail(User $user, string $setupUrl, int $expireMinutes): bool|string
    {
        try {
            Mail::to($user->email)->send(new UserCredentialsMail(
                user: $user,
                setupUrl: $setupUrl,
                expireMinutes: $expireMinutes,
                loginUrl: route('login'),
            ));

            ActivityLogger::log(
                action: 'mail_sent',
                description: "Sent account setup link to {$user->email}",
                subject: $user,
            );

            return true;
        } catch (\Throwable $e) {
            Log::warning('Failed to send user setup email', [
                'user_id' => $user->id,
                'email'   => $user->email,
                'error'   => $e->getMessage(),
            ]);

            return $e->getMessage();
        }
    }

    private function applyPermissions(Request $request, array &$data): void
    {
        foreach (array_keys(User::MODULES) as $module) {
            $editKey = "can_edit_{$module}";
            $viewKey = "can_view_{$module}";
            $edit = $request->boolean($editKey);
            $view = $request->boolean($viewKey) || $edit;
            $data[$viewKey] = $view;
            $data[$editKey] = $edit;
        }
    }
}
