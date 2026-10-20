module github.com/excelano/graphauth

go 1.26.0

// A bump to MSAL reaches every consumer on its next release. Read the MSAL
// changelog before taking a minor.
require (
	github.com/AzureAD/microsoft-authentication-library-for-go v1.10.1
	github.com/excelano/atrest v0.3.0
	golang.org/x/term v0.47.0
)

require (
	github.com/godbus/dbus/v5 v5.2.4 // indirect
	github.com/golang-jwt/jwt/v5 v5.3.1 // indirect
	github.com/google/uuid v1.6.0 // indirect
	github.com/kylelemons/godebug v1.1.0 // indirect
	github.com/pkg/browser v0.0.0-20240102092130-5ac0b6a4141c // indirect
	golang.org/x/sync v0.24.0 // indirect
	golang.org/x/sys v0.49.0 // indirect
)
