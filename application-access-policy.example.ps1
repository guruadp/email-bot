# Restrict each email-bot Azure app registration to ONLY its own mailbox,
# instead of the tenant-wide access "Mail.ReadWrite"/"Mail.Send" application
# permissions grant by default. Run this once from an account with Exchange
# Online admin rights. Requires the ExchangeOnlineManagement PowerShell module:
#
#   Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#
# Then run the commands below (PowerShell, not bash).

Connect-ExchangeOnline -UserPrincipalName admin@ednex.me

# --- Guru's app: dffe5a8e-ae87-433a-a053-d5afeb6cb308 -> guru@ednex.me only ---
New-DistributionGroup -Name "email-bot-guru-scope" -Type Security -Members guru@ednex.me
New-ApplicationAccessPolicy `
    -AppId "dffe5a8e-ae87-433a-a053-d5afeb6cb308" `
    -PolicyScopeGroupId "email-bot-guru-scope@ednex.me" `
    -AccessRight RestrictAccess `
    -Description "Email bot (Guru) - restrict to guru@ednex.me only"

# --- Senthil's app: ecc7a355-eff8-4014-8442-aa78d3bb0226 -> senthil@ednex.me only ---
New-DistributionGroup -Name "email-bot-senthil-scope" -Type Security -Members senthil@ednex.me
New-ApplicationAccessPolicy `
    -AppId "ecc7a355-eff8-4014-8442-aa78d3bb0226" `
    -PolicyScopeGroupId "email-bot-senthil-scope@ednex.me" `
    -AccessRight RestrictAccess `
    -Description "Email bot (Senthil) - restrict to senthil@ednex.me only"

# --- Verify: should say "Granted" for the owning mailbox... ---
Test-ApplicationAccessPolicy -AppId "dffe5a8e-ae87-433a-a053-d5afeb6cb308" -Identity guru@ednex.me
Test-ApplicationAccessPolicy -AppId "ecc7a355-eff8-4014-8442-aa78d3bb0226" -Identity senthil@ednex.me

# --- ...and "Denied" for every other mailbox in the tenant, e.g.: ---
Test-ApplicationAccessPolicy -AppId "dffe5a8e-ae87-433a-a053-d5afeb6cb308" -Identity senthil@ednex.me
Test-ApplicationAccessPolicy -AppId "ecc7a355-eff8-4014-8442-aa78d3bb0226" -Identity guru@ednex.me

# Policy changes can take up to 30 minutes to propagate.
