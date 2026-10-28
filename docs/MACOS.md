# Opening EDF Evidence Collector on a Mac

[Documentation index](README.md) · [Downloads](../README.md#download-a-build)

This download is for Apple silicon Macs (M1, M2, M3, M4 and later).
It uses an ad-hoc code signature and is not notarized by Apple. macOS will
normally ask you to approve the first launch of a downloaded copy.
An ad-hoc signature checks the app's integrity; it does not identify an
Apple-registered developer or make Gatekeeper trust the download automatically.

## First launch

1. Extract the downloaded zip. GitHub Actions downloads wrap the Mac zip inside
   another zip, so you may need to extract twice.
2. Inside the `EDF Evidence Collector` folder, drag
   `EDF_Evidence_Collector.app` to Applications. In Finder the `.app` suffix
   may be hidden; the app is the item with the application icon.
3. Open the app from Applications. If macOS blocks it, dismiss the alert.
4. Open Apple menu → System Settings → Privacy & Security. Scroll down to
   Security and click **Open Anyway** beside the message about this app.
   This option appears after a blocked launch and is available for about an hour.
5. Confirm with your Mac password or Touch ID, then click **Open** if asked.

Recent macOS versions do not put a Windows-style “continue anyway” button
on the initial warning. This approval is normally needed once for each
downloaded version. Only approve a copy obtained from this project's GitHub
repository: https://github.com/volster/edf-bill-fetcher

If the alert explicitly says the app **will damage your computer**, or your
Mac is managed and has no approval option, stop and record the exact wording.
That is a different condition from an unidentified developer warning.
Apple's instructions: https://support.apple.com/en-gb/102445

## What the download checks cover

The Mac build uses a standard `.app` bundle. CI verifies its nested code
signatures, creates the zip using Apple's `ditto`, extracts it again, and
verifies the extracted signatures. It then opens that extracted app through
macOS LaunchServices, requires its real Cocoa window to appear, and checks
bundled OCR and the Excel correction workflow.

`build-provenance.json` identifies the source revision and CI validation.
`macos-trust.json` records the architecture and Apple trust status.
These checks establish packaging and startup integrity on the build runner;
they do not substitute for Apple notarization or test the download's
quarantine approval on your Mac.

## Automatic Apple trust

Removing the first-launch approval requires a paid Apple Developer Program
membership, a Developer ID Application certificate, and Apple notarization.
The current workflows do not use those credentials. A future signed release
must sign all bundled native components during the PyInstaller build, submit
the app to Apple's notary service, staple the accepted ticket, and verify the
final extracted archive with Gatekeeper before describing it as notarized.
