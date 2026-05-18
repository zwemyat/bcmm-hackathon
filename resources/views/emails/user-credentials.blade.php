<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Welcome to {{ $appName }}</title>
</head>
<body style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
    <h2 style="color: #0d6efd; margin-bottom: .25rem;">Welcome to {{ $appName }}</h2>
    <p>Hello {{ $user->name }},</p>
    <p>An account has been created for you. To finish setting up, choose your own password using the secure link below:</p>

    <p style="margin: 24px 0;">
        <a href="{{ $setupUrl }}"
           style="display:inline-block; padding: 12px 22px; background: #0d6efd; color: #fff; text-decoration: none; border-radius: 6px; font-weight: 600;">
            Set your password
        </a>
    </p>

    <p style="color: #475569; font-size: 13px;">
        This link expires in <strong>{{ $expireMinutes }} minutes</strong> and can be used once.
        If it expires before you use it, please ask your administrator to issue a new one.
    </p>

    <table cellpadding="10" cellspacing="0" border="1" style="border-collapse: collapse; border-color: #e2e8f0; margin: 16px 0; font-size: 13px;">
        <tr>
            <td style="background:#f8fafc;"><strong>Your email</strong></td>
            <td>{{ $user->email }}</td>
        </tr>
        <tr>
            <td style="background:#f8fafc;"><strong>Role</strong></td>
            <td>{{ ucfirst($user->role) }}</td>
        </tr>
        <tr>
            <td style="background:#f8fafc;"><strong>Sign-in URL</strong></td>
            <td><a href="{{ $loginUrl }}">{{ $loginUrl }}</a></td>
        </tr>
    </table>

    <p style="color: #b91c1c; font-size: 13px;">
        <strong>Security tip:</strong> never share this link with anyone. Your administrator will never ask you for your password.
    </p>

    <p style="color: #6b7280; font-size: 12px;">— {{ $appName }}</p>
</body>
</html>
