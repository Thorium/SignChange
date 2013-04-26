Outlook plug-in to insert a random quote to signature.

Run: SignatureChange\bin\Release\SignatureChange.vsto

Changes every 5 minutes and when Outlook restarts.

First insert a signature in Outlook:
File -> Options -> Mail -> Signatures
Ensure that you have "..." in the signature, place where the quote should be added.
(Annoyingly Outlook tries to replace "..." with

This will replace the three dots "..." location from the Signature file with a random quote.

Reads signatures from: SignatureChange\Properties\Resources.resx
