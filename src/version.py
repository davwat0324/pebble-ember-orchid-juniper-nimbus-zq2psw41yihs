"""The app's version and where its releases are published.

release.ps1 bumps __version__; the GitHub Actions workflow builds a release whose tag is v<__version__>.
"""
__version__ = "1.1.1"

# Public GitHub repository whose Releases carry Remind-win64.zip (the updater reads the latest one).
REPO = "davwat0324/pebble-ember-orchid-juniper-nimbus-zq2psw41yihs"
ASSET = "Remind-win64.zip"
