# aras-vendor-sync

Scheduled PowerShell service that refreshes per-vendor pricing on
`Vendor Part` rows in Aras Innovator 2025. Runs as a Windows Scheduled
Task on the Aras VM. Independent of the [aras-mcp-server](https://github.com/Tungsten-lawrenceb/aras-mcp-server)
MCP — that MCP captures pricing the first time a datasheet is fetched;
this service keeps the data fresh on a schedule.

## What it does

Once per run (weekly Sundays at 03:00 by default):

1. OAuth into Aras as the `aras-vendor-sync` service identity.
2. Read `MPN_API_Config` rows from Aras for vendor credentials
   (Digi-Key client_id/client_secret, Mouser API key).
3. List every `Vendor Part` row that carries a `catalog_number`.
4. For each row, call the vendor's API to fetch fresh pricing.
5. PATCH `unit_price`, `currency`, `min_order_qty`, `availability`,
   `product_status`, `pricing_updated_on` on the Vendor Part row.
6. `MCP_RecomputeMPUnitPrice` (a server-side Method bound to Vendor
   Part's `onAfterUpdate` event) automatically rolls each updated
   Vendor Part price into its parent `Manufacturer Part.unit_price`
   as the min across all vendors.

Datasheets and Vendor Part creation are owned by the MCP, not this
service. This script only refreshes pricing on rows that already exist.

## Why a separate repo from aras-mcp-server

The MCP is a developer-facing Python tool for ad-hoc PLM work. This is
unattended production infrastructure on the Aras VM with a different
operational profile: runs without the MCP, has its own service account,
its own credentials, its own logs, its own deployment cadence. Bundling
them would conflate two very different lifecycles.

## Install

One-time, on the Aras VM as Administrator:

```powershell
# 1. Drop this repo under C:\Share\customizations\src\
git clone https://github.com/Tungsten-lawrenceb/aras-vendor-sync.git `
    C:\Share\customizations\src\aras-vendor-sync

# 2. Copy and fill the config template
$cfg = 'C:\ProgramData\AarasVendorSync\config.json'
New-Item -ItemType Directory -Path (Split-Path $cfg) -Force | Out-Null
Copy-Item C:\Share\customizations\src\aras-vendor-sync\config\config.example.json $cfg
notepad $cfg   # fill in aras user/password_md5; vendor creds come from MPN_API_Config

# 3. Tighten the ACL on the config file (Administrators + the service
#    identity only)
icacls $cfg /inheritance:r /grant 'Administrators:F' "$env:USERNAME:R"

# 4. Register the Scheduled Task (weekly Sundays 03:00 by default)
& C:\Share\customizations\src\aras-vendor-sync\scripts\Install-ScheduledTask.ps1
```

Manual one-off run for testing:

```powershell
& C:\Share\customizations\src\aras-vendor-sync\src\Refresh-VendorPricing.ps1 -ConfigPath $cfg -DryRun
```

`-DryRun` reports what would be patched without writing anything. Drop
the flag to apply.

## Service account

Best practice is a dedicated `aras-vendor-sync` Aras user, not
`mcp-service`. Steps:

```powershell
# In Aras (UI or MCP):
# - Create User "aras-vendor-sync"
# - Add to "Innovator Admin" identity group (needed to PATCH Vendor Part)
# - Generate a password, MD5-hash it, store the hex in config.json
```

The MCP can do these steps via `create_user` + `add_member` if you'd
rather scripted.

## Logs

Each run appends one line per Vendor Part processed to
`C:\ProgramData\AarasVendorSync\logs\YYYY-MM-DD.log` plus a summary at
the end (counts of attempted / succeeded / failed / skipped). Errors
also write to the Windows Application Event Log under source
`AarasVendorSync`.

## Quota

The default cadence (weekly) plus the size of the Tungsten BoM (~72
MPs) is ~72 Digi-Key ProductDetails calls per Sunday: ~5% of the daily
free-tier budget. Daily refresh would be 30× that — still inside the
1000/day free tier. Adjust the trigger in
`scripts/Install-ScheduledTask.ps1` if you need different cadence.
