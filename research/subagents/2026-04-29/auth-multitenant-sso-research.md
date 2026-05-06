# Research: Multi-tenant Auth for an Outlook Add-in calling the CCM SaaS API

Date: 2026-04-29
Topic: Office SSO + multi-tenant Entra ID + OBO + fallback dialog for "Contoso Consents Management" (CCM) Outlook add-in.

## Scope and questions

1. How does Office SSO (`Office.auth.getAccessToken` / `OfficeRuntime.auth.getAccessToken`) work for an Outlook add-in, what token does it return, what are the manifest and app registration prerequisites, and how does on-behalf-of (OBO) extend it to call the CCM API?
2. For a multi-tenant SaaS (CCM) sold to many customer tenants: should the Outlook add-in front-end and the CCM API share one app registration or use two? What are the well-known Office host client IDs that must be preauthorized?
3. How does tenant admin consent at customer onboarding work (`/adminconsent` endpoint)?
4. What is the recommended fallback when SSO is unavailable (Outlook mobile, certain tenants, guest users, MFA / conditional access blocking SSO)? Office Dialog API + MSAL.js auth code + PKCE; reference the official `Office-Add-in-samples` fallback sample.
5. How should the CCM API validate bearer tokens given a multi-tenant issuer and a `tid` allow-list of paying customer tenants?
6. What conditional access / continuous access evaluation (CAE) considerations apply?

## Status: Complete

All primary questions answered with authoritative Microsoft Learn sources. One important addition: Microsoft Learn now flags the classic SSO+OBO flow as "legacy" and recommends MSAL.js Nested App Authentication (NAA) as the modern replacement. Both patterns are documented below; the user-requested legacy pattern is the primary focus.

---

## 1. Office SSO for Outlook add-ins

### Token returned by `getAccessToken`

- API: `OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true })` (preferred) or `Office.auth.getAccessToken`. Requires the `IdentityAPI 1.3` requirement set.
- The returned string is a **JWT issued by the Microsoft identity platform** to the add-in's own Entra ID app registration. It is simultaneously an **identity token** (carries `name`, `preferred_username`, `oid`, `tid`) and an **access token** scoped to the add-in's API (default scope `access_as_user`).
- Typical decoded payload (from Microsoft Learn `sso-in-office-add-ins`):
  - `aud` = the add-in's app registration client ID (or its `api://...` URI for v1.0 tokens)
  - `iss` = `https://login.microsoftonline.com/<tenantId>/v2.0`
  - `scp` = `access_as_user`
  - `tid` = the signed-in user's home tenant ID
  - `oid`, `preferred_username`, `name` — user identity claims
- Office caches the token; never cache it yourself. Always call `getAccessToken` again when needed.
- Source: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins>

### Manifest prerequisites

Both add-in-only XML and unified JSON manifests need a `webApplicationInfo` block linking the manifest to the Entra app registration:

Unified JSON manifest:
```json
"webApplicationInfo": {
  "id": "<add-in-app-registration-client-id-guid>",
  "resource": "api://addin.contoso.com/<add-in-app-registration-client-id-guid>"
}
```

XML add-in-only manifest equivalent: `<WebApplicationInfo><Id>...</Id><Resource>api://...</Resource><Scopes><Scope>profile</Scope><Scope>openid</Scope></Scopes></WebApplicationInfo>`.

The `resource` value must match the Application ID URI configured under "Expose an API" and must end with the client ID GUID.

Source: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins>

### Entra app registration prerequisites (legacy SSO)

From <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2>:

1. Register an app, **Supported account types: "Accounts in any organizational directory (any Microsoft Entra directory - multitenant) and personal Microsoft accounts"** for a SaaS sold to many tenants.
2. Add an SPA redirect URI for the fallback dialog (e.g. `https://addin.contoso.com/dialog.html`).
3. Add a client secret (server keeps this for the OBO call).
4. Expose an API → set Application ID URI to `api://<fully-qualified-domain-name>/<app-id>` (e.g. `api://addin.contoso.com/c6c1f32b-...`).
5. Add scope **`access_as_user`** (state Enabled, "Admins and users" can consent).
6. **Pre-authorize Office host client IDs** under the new scope so Office can silently obtain a token for the add-in without per-user consent on those clients.
7. Add the delegated Microsoft Graph permissions the add-in needs (e.g. `profile`, `openid`, `Mail.Read`, `Files.ReadWrite`); grant tenant admin consent for the publisher's home tenant.
8. In the app manifest, set `requestedAccessTokenVersion` = `2` to receive v2.0 tokens with `iss = https://login.microsoftonline.com/{tid}/v2.0`.

### Office host client IDs to preauthorize

Microsoft Learn (`register-sso-add-in-aad-v2`) lists these as the trusted Office client app IDs:

| Client ID | Host |
|---|---|
| `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` | Microsoft Office (covers all Office endpoints; preferred single value) |
| `d3590ed6-52b3-4102-aeff-aad2292ab01c` | Microsoft Office (desktop) |
| `93d53678-613d-4013-afc1-62e9e444a0a5` | Office on the web |
| `bc59ab01-8403-45c6-8796-ac3ef710b3e3` | Outlook on the web |
| `1fec8e78-bce4-4aaf-ab1b-5451cc387264` | Microsoft Teams desktop / Teams mobile |
| `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` | Teams on the web |

Recommendation: use `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` to cover all platforms with one preauthorized application entry; add the Teams IDs only if the add-in must also work in Teams. The user-supplied IDs `d3590ed6...` (Microsoft Office) and `4345a7b9-9a63-4910-a426-35363201d503` are also referenced in older Microsoft samples; only the IDs published in `register-sso-add-in-aad-v2` are authoritative for Office add-in SSO today, so prefer the table above.

Note on the user's request: the value `4345a7b9-9a63-4910-a426-35363201d503` does not appear in the current Microsoft Learn SSO registration article. It is the well-known Outlook Mobile (`com.microsoft.Office.Outlook` iOS/Android) public client ID seen in some samples but is **not required** for Office add-in SSO; preauthorization is at the Office host level (`ea5a67f6-...`), not the Outlook native app level. Outlook mobile add-ins should rely on the fallback dialog flow instead (see §4).

### On-Behalf-Of (OBO) flow to call CCM API

If the CCM API is a separate app registration from the add-in, the add-in's server-side code exchanges the SSO token for a token to CCM. If the add-in front-end and CCM API share one registration (recommended — see §2), OBO is only needed to call **further** downstream APIs (e.g. Microsoft Graph). When the audience of the SSO token already equals the CCM API, OBO is not needed for the call from the add-in's middle-tier to CCM itself; the SSO token is the API token.

OBO request (server-side, RFC-shaped form post to `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`):

```
grant_type        = urn:ietf:params:oauth:grant-type:jwt-bearer
client_id         = <add-in app registration client id>
client_secret     = <secret>      (or client_assertion + client_assertion_type for cert auth)
assertion         = <SSO token from getAccessToken>
scope             = https://graph.microsoft.com/Mail.Read offline_access
                    (or api://ccm.contoso.com/<ccm-api-app-id>/.default)
requested_token_use = on_behalf_of
```

Use `{tenant}` = the user's `tid` from the SSO token (not `/common`) to avoid MSAL caching pitfalls in multi-tenant scenarios.

Important multi-tenant OBO requirements:
- The middle-tier app must declare the front-end client in `knownClientApplications` if they are separate registrations, so combined consent surfaces both apps' permissions in one screen.
- Use `/.default` scope to acquire the token using statically configured permissions and combined consent.
- Do **not** mix `/.default` with other delegated scopes in the same OBO call (causes `AADSTS70011`).
- Never return the OBO-acquired downstream token to the client.

Source: <https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow>

---

## 2. App registration design: one or two registrations?

### Microsoft Learn guidance

From the OBO doc, "Use of a single application":
> "In some scenarios, you could only have a single pairing of middle-tier and front-end client. In this scenario, you could find it easier to make this a single application, negating the need for a middle-tier application altogether."

### Recommendation for CCM

**Use one multi-tenant app registration** that simultaneously:
- Represents the Outlook add-in (has the SPA redirect URI for the fallback dialog).
- **Exposes the CCM API** under `api://ccm.contoso.com/<app-id>` with the scope `access_as_user` (or `CCM.Access` etc.).
- Preauthorizes Office host client IDs against that scope.
- Holds delegated permissions for any downstream APIs (Microsoft Graph) the CCM API will call via OBO.

Why one registration:
- The Office SSO token's `aud` is the add-in's app. If CCM is the same app, that token is already a valid bearer for CCM; no extra OBO call from the add-in's web frontend → CCM is needed.
- Simpler tenant admin consent: one consent grant covers add-in sign-in + CCM API access + downstream Graph permissions in a single screen.
- Eliminates `knownClientApplications` plumbing and a class of "AADSTS65001 – consent not granted" errors.

When to use two registrations:
- CCM is also consumed by non-Office clients (web SPA, mobile app, third-party integrations) where you need a different identity surface, lifecycle, or secret rotation policy.
- You need stricter app-id-based policy separation (e.g. different conditional access for the add-in vs. the API).
- In that case the add-in app declares `knownClientApplications: [<addin-app-id>]` on the CCM API app and uses the standard OBO flow.

Multi-tenant registration must:
- Set Supported account types to "Accounts in any organizational directory" (or include personal accounts if needed).
- Have a globally unique Application ID URI tied to a verified custom domain (e.g. `api://ccm.contoso.com/<app-id>`). Publisher domain should be verified to avoid the unverified-publisher consent warning.

Source: <https://learn.microsoft.com/en-us/entra/identity-platform/howto-convert-app-to-be-multi-tenant>

---

## 3. Tenant admin consent at customer onboarding

### `/adminconsent` endpoint

Per <https://learn.microsoft.com/en-us/entra/identity-platform/v2-admin-consent>:

```
GET https://login.microsoftonline.com/{tenant}/v2.0/adminconsent
    ?client_id=<ccm-app-id>
    &scope=https://graph.microsoft.com/.default
    &redirect_uri=https://app.ccm.contoso.com/onboarding/consent-complete
    &state=<opaque-state>
```

- `tenant`: customer's tenant ID or domain (e.g. `contosocustomer.onmicrosoft.com`). Do **not** use `common` for admin consent.
- A Global Administrator (or Privileged Role Administrator / Cloud Application Administrator) of the customer tenant must sign in and approve.
- After success, Entra redirects to `redirect_uri` with `?admin_consent=True&tenant=<tid>&state=...&scope=...`.
- Once granted, a service principal for CCM is created in the customer tenant with the requested permissions consented for **all users**, eliminating per-user consent prompts.

### Onboarding flow for CCM

1. Customer admin signs up at `https://app.ccm.contoso.com/onboarding`.
2. CCM signs the admin in (OIDC) so it knows the tenant ID.
3. CCM presents a "Grant CCM access for your organization" button → redirects to `/adminconsent` for that tenant.
4. On `admin_consent=True` callback, CCM stores the customer tenant ID in its **allow-list of paying tenants**.
5. End users in that tenant can now use the Outlook add-in without further consent prompts.

### Security warning from Microsoft Learn

> "Never use the tenant ID value of the `tenant` parameter to authenticate or authorize users. The tenant ID value can be updated and sent by bad actors to impersonate a response to your app."

Use `state` (cryptographically random, server-stored) to correlate the redirect; only trust `tid` claims from validated tokens, not the `tenant` query parameter.

---

## 4. Fallback when SSO fails or is unavailable

### When fallback is needed

`getAccessToken` can fail or be unsupported on:
- Outlook mobile (Outlook for iOS / Android) — full SSO not always supported.
- Older Office clients without IdentityAPI 1.3.
- Guest users in the host tenant (their home tenant differs from the host).
- Tenants with conditional access policies or MFA requirements that the silent SSO path can't satisfy (token returned with `interaction_required`-class errors).
- Domain-joined accounts that aren't Microsoft 365 work/school accounts (error code `13003`).
- Browsers blocking third-party cookies (Safari ITP, Chrome storage partitioning).

Reference: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins> "Implement a fallback authentication system".

### Fallback architecture: Office Dialog API + MSAL.js auth code + PKCE

Per <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api>:

1. The task pane catches the `getAccessToken` failure and calls `Office.context.ui.displayDialogAsync('https://addin.contoso.com/dialog.html', { height: 60, width: 30 })`.
2. The dialog is a **separate browser/webview instance**, not an iframe — so Entra ID's sign-in page (which refuses to render in iframes via `X-Frame-Options`) loads cleanly.
3. The dialog page uses MSAL.js (browser SDK) **authorization code flow with PKCE** against the same multi-tenant app registration. Redirect URI is the SPA redirect on `dialog.html`.
4. After MSAL acquires the token, the dialog calls `Office.context.ui.messageParent(JSON.stringify({ accessToken, expiresOn }))`.
5. The task pane receives the message via the `DialogMessageReceived` event handler, closes the dialog, and uses the token to call CCM.

### Token storage caveat

Per Microsoft Learn: the dialog and task pane are separate browser instances, so MSAL's in-memory cache in the dialog is **not visible** to the task pane. Strategies:
- Pass the token via `messageParent` (string only).
- For silent renewal, repeat the dialog pop or store the token server-side keyed by the user identity.
- `localStorage` is unreliable across runtimes (Safari, Chrome storage partitioning since v115, Edge default partitioning).

### Reference sample

The user mentioned "fallbackauthdialog". The current authoritative samples in `OfficeDev/Office-Add-in-samples`:

- `Samples/auth/Office-Add-in-NodeJS-SSO` — Node.js SSO with MSAL.js fallback dialog (this is the modern home of the original "fallbackauthdialog" sample structure).
- `Samples/auth/Office-Add-in-ASPNET-SSO` — ASP.NET SSO with fallback dialog.
- `Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET` — Outlook-specific MSAL.NET + dialog example.
- `Samples/auth/Office-Add-in-Microsoft-Graph-React` — React + msal.js inside the dialog.

(Note: the specific path `Samples/auth/Outlook-Auth-MSAL-NAA` returned 404 at the time of research. The NAA-based modern Outlook sample lives under another path; verify in the repo's `Samples/auth/` index.)

### Modern alternative: Nested App Authentication (NAA)

Microsoft Learn now marks classic SSO+OBO+fallback dialog as **legacy** and points to **Nested App Authentication (NAA)** with MSAL.js as the recommended modern path. Reference: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in>. NAA:
- Uses MSAL.js directly inside the task pane (no dialog needed for SSO).
- Returns tokens for any resource the app is consented to (eliminates the OBO server hop).
- Works on Outlook desktop, web, and mobile (where supported).
- Still benefits from the same multi-tenant app registration design described in §2.
- Recommended for new add-ins; legacy SSO+OBO+dialog remains supported for existing implementations.

Recommendation for CCM: build NAA as the primary path with the Office Dialog + MSAL.js fallback only for hosts that don't yet support NAA.

---

## 5. Securely passing tokens to and validating tokens at the CCM API

### Sending the token

```http
GET /api/consents
Host: api.ccm.contoso.com
Authorization: Bearer eyJ0eXAiOiJKV1Qi...
```

The CCM client sends a v2.0 access token whose `aud` is the CCM API. Use HTTPS only; never log tokens.

### Validating the bearer token at CCM (multi-tenant)

Reference: <https://learn.microsoft.com/en-us/entra/identity-platform/access-tokens#validate-tokens>.

For a multi-tenant API the validation is more involved than single-tenant:

1. **Configure the metadata endpoint as tenant-independent**: `https://login.microsoftonline.com/organizations/v2.0/.well-known/openid-configuration` (or `/common/...` if personal MSAs are accepted).
2. **Validate signature** using the JWKS at `jwks_uri`, selecting the key by `kid`. Each key's metadata exposes an `issuer` field — keys with templated issuer `https://login.microsoftonline.com/{tenantid}/v2.0` may sign tokens from any tenant; keys with concrete tenant GUID issuer (e.g. consumers `9188040d-...`) only sign for that tenant.
3. **Validate `iss`**: replace `{tenantid}` in the metadata issuer template with the token's `tid` claim and require exact match. Equivalently, require `iss == "https://login.microsoftonline.com/" + tid + "/v2.0"`.
4. **Validate `tid` is a GUID** and matches the issuer URL path segment.
5. **Validate `aud`**: must equal the CCM API's app id (or the `api://...` URI for v1.0 tokens). Setting `requestedAccessTokenVersion=2` forces v2 tokens where `aud` is the GUID.
6. **Validate `scp`**: must contain `access_as_user` (or whichever scope CCM requires).
7. **Validate `exp`, `nbf`, `iat`** with normal clock skew (~5 min).
8. **Tenant allow-list (CCM-specific)**: after issuer validation passes, check `tid ∈ AllowedCustomerTenants` (the set of customers who completed the `/adminconsent` flow). Reject with HTTP 403 + a clear "tenant not provisioned" error if not.
9. **Per-tenant data scoping**: always include `tid` in the lookup key for any tenant-owned data — Microsoft Learn explicitly says claims like `sub` and `oid` must be interpreted within the issuing tenant.

Recommended libraries:
- .NET: `Microsoft.Identity.Web` — set `AzureAd:TenantId = "organizations"` and use the `TokenValidatedEvent` to enforce the `tid` allow-list.
- Node.js: `passport-azure-ad` or `jwt-rsa` + `jose` with `issuer = (token) => https://login.microsoftonline.com/${token.tid}/v2.0`.
- Python: `msal` for clients; `PyJWT` + JWKS client for resource servers.

### Common pitfalls

- Validating `iss` against a fixed string — fails for any tenant other than the publisher's.
- Trusting unverified `tid` from the request body or query string. Always derive `tid` from the validated token.
- Accepting tokens whose `kid`'s issuer is the consumer tenant when CCM doesn't support MSAs.
- Allowing wildcard `aud` — must be the CCM app ID exactly.

---

## 6. Conditional Access and Continuous Access Evaluation

Reference: <https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-continuous-access-evaluation>.

### What matters for CCM

- A customer tenant may apply Conditional Access policies (MFA, location, device compliance, sign-in frequency) to the add-in app, even though CCM's home tenant doesn't define them. Honor `interaction_required`, `consent_required`, and `claims` challenges returned by Entra during OBO; surface them to the client as **HTTP 401 with a `WWW-Authenticate: Bearer error="insufficient_claims", claims="<base64>"` header** so the client can re-acquire a token with the claims challenge.
- The OBO doc explicitly shows the error JSON format with a `claims` field that must be propagated.
- CAE is currently focused on Exchange Online, SharePoint Online, Teams, MS Graph. A custom SaaS API like CCM is not a CAE-enabled resource by default, but to participate (and benefit from up-to-28-hour token lifetimes plus near-real-time revocation) the API can opt in by:
  1. Declaring CAE capability when issuing tokens (the resource declares `xms_cc` capability).
  2. Implementing the **claim-challenge protocol** (return 401 with `WWW-Authenticate` containing the claims challenge when access should be revoked or re-evaluated).
- See <https://learn.microsoft.com/en-us/entra/identity-platform/app-resilience-continuous-access-evaluation> and <https://learn.microsoft.com/en-us/entra/identity-platform/claims-challenge>.

### Practical implications

- Office (Word/Excel/PowerPoint/Outlook desktop+web) supports claim challenges; the add-in's token-acquisition code must handle 401+claim-challenge by calling `getAccessToken({ authChallenge: <claims> })` (or the MSAL.js fallback with claims) so Entra re-authenticates the user against the policy.
- The OBO middle tier should never retry with a cached token after a 401-claims response.
- Guest users: CAE doesn't enforce instant revocation for guest accounts — document this limitation.
- Sign-in Frequency policies are honored regardless of CAE; tokens may have lifetimes from minutes to hours depending on the customer's CA configuration.

---

## Decision summary for CCM

| Decision | Recommended choice |
|---|---|
| App registration | One multi-tenant Entra app for both add-in front-end and CCM API |
| Account types | "Accounts in any organizational directory" (add personal MSA only if needed) |
| Application ID URI | `api://ccm.contoso.com/<app-id>` (verified publisher domain) |
| Exposed scope | `access_as_user` (or domain-specific name like `Consents.Access`) |
| Preauthorized clients | `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (covers all Office hosts), plus Teams IDs if needed |
| `requestedAccessTokenVersion` | `2` |
| Manifest | `webApplicationInfo` with id+resource pointing at the registration |
| Primary auth path | NAA (MSAL.js + nested app auth) on supported hosts |
| Fallback auth path | Office Dialog API + MSAL.js auth code + PKCE on `dialog.html` |
| OBO | Only when calling downstream APIs (Microsoft Graph etc.), not for the CCM API itself |
| Customer onboarding | `/adminconsent` flow, persist customer `tid` to allow-list |
| Token validation | Multi-tenant: validate `iss == https://login.microsoftonline.com/{tid}/v2.0`, `aud == ccm-app-id`, `scp` contains required scope, `tid` ∈ allow-list |
| CAE | Implement claim-challenge propagation (401 + WWW-Authenticate) end-to-end |

---

## Authoritative references

- Office SSO overview: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins>
- Outlook SSO authentication article: <https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authenticate-a-user-with-an-sso-token>
- Register SSO add-in (Office host client IDs): <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2>
- Office Dialog API for auth: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api>
- Nested App Authentication (modern alternative): <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in>
- OAuth 2.0 On-Behalf-Of flow: <https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow>
- Convert single-tenant app to multitenant: <https://learn.microsoft.com/en-us/entra/identity-platform/howto-convert-app-to-be-multi-tenant>
- Admin consent endpoint: <https://learn.microsoft.com/en-us/entra/identity-platform/v2-admin-consent>
- Access tokens (multi-tenant validation): <https://learn.microsoft.com/en-us/entra/identity-platform/access-tokens>
- Continuous Access Evaluation: <https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-continuous-access-evaluation>
- CAE-enabled APIs guidance: <https://learn.microsoft.com/en-us/entra/identity-platform/app-resilience-continuous-access-evaluation>
- Claims challenges: <https://learn.microsoft.com/en-us/entra/identity-platform/claims-challenge>
- Office Add-in samples (auth folder root): <https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth>
- Node.js SSO sample (closest current equivalent of "fallbackauthdialog"): <https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO>
- ASP.NET SSO sample: <https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO>
- Outlook Microsoft Graph ASP.NET sample: <https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET>

---

## Recommended next research (not completed)

- [ ] Verify the exact current path of the NAA-based Outlook sample (the URL `Samples/auth/Outlook-Auth-MSAL-NAA` returned 404; locate the renamed/moved sample).
- [ ] Concrete `xms_cc` capability and `cnf` claim wiring needed for the CCM API to be CAE-enabled (vs. just claim-challenge aware).
- [ ] Recommended pattern for tenant-admin off-boarding (revoking `tid` from CCM allow-list) and how to detect deletion via Microsoft Graph subscription / `tenantInformation`.
- [ ] Specific MSAL.js v3 / v4 NAA bootstrap code for an Outlook add-in (sample fragment).
- [ ] Whether personal MSAs (Outlook.com) should be supported and the implication for the `aud` / issuer validation logic (`organizations` vs. `common`).

## Clarifying questions for the user

1. Will CCM ever be called by clients other than the Outlook add-in (web SPA, mobile app, server-to-server)? If yes, two app registrations (with proper `knownClientApplications`) become more attractive.
2. Should personal Microsoft accounts (Outlook.com) be supported, or strictly work/school accounts? This drives `/common` vs. `/organizations` and whether the consumers tenant `9188040d-6c67-4c5b-b112-36a304b66dad` keys must be accepted.
3. Does CCM need to call Microsoft Graph (or any other Entra-protected API) on behalf of the user? If not, the OBO flow can be eliminated entirely with a single app registration.
4. Are you targeting Outlook mobile (iOS/Android) at GA, or only Outlook desktop and web initially? This affects whether NAA suffices or the legacy fallback dialog must also ship.
5. Is the CCM API hosted under a verified publisher domain in Entra ID? Required to avoid the unverified-publisher consent warning at customer onboarding.
6. What conditional access posture do target customers typically run (MFA, sign-in frequency, device compliance)? Drives priority of CAE / claim-challenge implementation.
