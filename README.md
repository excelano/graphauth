# graphauth

The Microsoft Graph sign-in layer shared by Excelano's command-line tools, among them [xql](https://github.com/excelano/xql) and [xfiles](https://github.com/excelano/xfiles). It carries the device-code OAuth flow over MSAL, a sealed on-disk token cache, the refusal that keeps an unattended caller from hanging on a code nobody will read, a table of AADSTS error hints, and a thin authenticated Microsoft Graph HTTP client.

It is a library for those tools, not a general Graph SDK. Everything hangs off an `App`: one Entra app registration, the delegated scopes its tools request, the name of its token cache, and the tools that sign in against it. Tools on the same `App` share consent and one cached session; tools on different `App`s never touch each other's. `graphauth.SharePoint` is the "Excelano SharePoint tools" registration; a tool on another registration declares its own `App`.

## Install

```
go get github.com/excelano/graphauth
```

## Usage

```go
import "github.com/excelano/graphauth"

app := graphauth.SharePoint

client, err := app.NewPublicClient(app.CachePath(), legacyCachePath)
if err != nil { /* setup failure */ }

result, err := app.Authenticate(ctx, client)
if err != nil {
    fmt.Fprintf(os.Stderr, "authentication failed: %v%s\n", err, graphauth.HintForAuthError(err))
    os.Exit(1)
}

graph := app.NewGraphClient(client, result.Account)
body, err := graph.Get(ctx, "/sites/"+siteID+"/lists", nil)
```

`Authenticate` tries a silent refresh against the cached account first and falls back to device code, printing the code and URL on stderr. When that fallback would be needed and stderr is not a terminal it returns `ErrNoTerminal` at once instead of polling for fifteen minutes: the remedy is in the message, and a cached refresh token still renews unattended, so the guard only ever bites the first sign-in.

`NewGraphClient` takes options. `WithTimeout` bounds each request; the default of five minutes is sized for file content. `WithHeader` adds a header to every authenticated request, which is how xql sends the `Prefer` header SharePoint needs before it will `$filter` on non-indexed list columns.

Resources with optimistic concurrency, Planner's among them, refuse a write that does not name the version it replaces. `PatchIfMatch` and `DeleteIfMatch` send the resource's `@odata.etag` as `If-Match`. A non-2xx answer comes back as an `*HTTPError`, so a caller can reach the status with `errors.As` and tell a stale etag (412) from a missing resource (404).

## The token cache

`App.CachePath` is `~/.config/excelano/<Cache>.json` (under `$XDG_CONFIG_HOME` when that is set), one file per `App`, so signing in with any of its tools signs in all of them. The second argument to `NewPublicClient` names a per-tool cache a consumer kept before its cache was shared; when the shared file does not exist yet and that one does, it is copied into place on first use and nobody signs in again. Where the cache is sealed, the legacy file is then deleted, since it would be a plaintext copy of the same token. Pass `""` for a consumer that never had one.

The file holds MSAL's own format, sealed through [atrest](https://github.com/excelano/atrest) with the operating system's data protection where the platform has one, and plaintext where it does not. A plaintext cache is read as it is and stored sealed on the same read, and a build that cannot open a sealed cache reads it as empty and asks for one sign-in rather than failing. The file is written 0600 in a 0700 directory through a temp file and a rename, so a crash or a second process writing the same file leaves the previous cache intact rather than a truncated one.

## Testing against a tenant

`go test ./...` covers the cache, the refusal, the hints and the state command without a token. The Graph client is covered by a live suite behind a build tag, which runs on a machine that has signed in to the SharePoint registration once and needs a test site to work in:

```
GRAPHAUTH_LIVE_SITE=https://<tenant>.sharepoint.com/sites/<test-site> go test -tags live ./...
```

Every request shape the client offers runs once against that site's default library, in a folder the test creates and removes. It does not run in CI: device-code sign-in needs a human, and a refresh token in a public repository's secrets would be a live credential to the tenant.

## The state command

Every consuming tool answers a bare `auth` with the state of its `App`'s session — account, tenant, token expiry and scopes — and exits 0 whether or not one exists, so a caller that would rather ask than find out can branch on the report instead of on a failed attempt. `App.AuthCommand` is that subcommand, flag parsing and rendering included, so the tools agree by construction; `App.CheckStatus` is the underlying question for a program that wants the `Status` value. Neither starts a sign-in: with no cached account the answer is offline, and with one it is a silent token renewal, which proves the refresh token still works. `--json` prints the same fields as an object.

## Not a consumer

blick-cli hand-rolls `x/oauth2` against per-tenant mailbox scopes. That is a different design on purpose and should not be folded in here.

## License

MIT.
