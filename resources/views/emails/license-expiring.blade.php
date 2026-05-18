<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>License Expiry Reminder</title>
</head>
<body style="font-family: Arial, sans-serif; color: #333;">
    @if(!empty($isTest))
        <div style="padding: 12px 16px; background: #fff7ed; border: 1px solid #fdba74; border-radius: 6px; color: #9a3412; margin-bottom: 20px; font-size: 13px;">
            <strong>TEST EMAIL</strong> &mdash; This is a preview of the license expiry reminder format using a real license from the system. No action is required; this message was triggered manually from Mail Settings.
        </div>
    @endif

    <h2 style="color: #d9534f;">License Expiry Reminder</h2>
    <p>Hello Admin,</p>
    <p>The following license is expiring in <strong>{{ $daysRemaining }} day(s)</strong>:</p>
    <table cellpadding="8" cellspacing="0" border="1" style="border-collapse: collapse; border-color: #ccc;">
        <tr><td><strong>Software</strong></td><td>{{ $license->software_name }}</td></tr>
        <tr><td><strong>Vendor</strong></td><td>{{ $license->vendor_name }}</td></tr>
        <tr><td><strong>License Info</strong></td><td>{{ $license->license_info }}</td></tr>
        <tr><td><strong>Expire Date</strong></td><td>{{ optional($license->expire_date)->format('Y-m-d') }}</td></tr>
        <tr><td><strong>Last Renewal Date</strong></td><td>{{ optional($license->last_renewal_date)->format('Y-m-d') ?: '—' }}</td></tr>
        <tr><td><strong>Renewal Type</strong></td><td>{{ $license->renewal_type }}</td></tr>
        <tr><td><strong>Renewal Cost</strong></td><td>{{ $license->renewal_cost }} {{ $license->currency }}</td></tr>
        <tr><td><strong>Status</strong></td><td>{{ $license->status }}</td></tr>
    </table>
    <p>Please take action to renew or terminate this license before it expires.</p>
    <p>— ITAMS Notification System</p>
</body>
</html>
