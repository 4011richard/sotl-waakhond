/* SOTL Waakhond: bij Verzenden van een mail naar buiten zonder #nummer of #info in het onderwerp
   toont Outlook een venster "Niet verzenden / Toch verzenden". Leest enkel onderwerp en ontvangers, stuurt niets door. */
var INTERN = ["step-out.be", "3broers.onmicrosoft.com", "b-4.be", "leefbaarcomfort.be", "comfortdienstencheques.be", "comfortprof.be"];
var CODES = ["#info"];          // uit brein/outlook/waakhond.md
var VRIJ = [];                  // ontvangers (adres of @domein) die geen #nummer nodig hebben, uit waakhond.md
var REF = /#\s?\d{2,5}(?!\d)/;

function eindigtOp(adres, dom) {
  return adres === dom || adres.slice(-(dom.length + 1)) === "@" + dom || adres.slice(-(dom.length + 1)) === "." + dom;
}

function nodigNummer(adres) {
  adres = (adres || "").toLowerCase();
  if (!adres || adres.indexOf("@") < 0) return false;
  for (var i = 0; i < INTERN.length; i++) if (eindigtOp(adres, INTERN[i])) return false;
  for (var j = 0; j < VRIJ.length; j++) {
    var v = VRIJ[j].toLowerCase();
    if (v.charAt(0) === "@" ? eindigtOp(adres, v.slice(1)) : adres === v) return false;
  }
  return true;
}

function lees(veld) {
  return new Promise(function (ok) {
    var item = Office.context.mailbox.item;
    var leeg = (veld === "subject" || veld === "body") ? "" : [];
    if (!item[veld]) return ok(leeg);
    var klaar = function (r) { ok(r.status === Office.AsyncResultStatus.Succeeded ? r.value : leeg); };
    if (veld === "body") item.body.getAsync(Office.CoercionType.Text, klaar);
    else item[veld].getAsync(klaar);
  });
}

function onMessageSendHandler(event) {
  Promise.all([lees("subject"), lees("to"), lees("cc"), lees("bcc"), lees("body")]).then(function (r) {
    var tekst = (r[4] || "").split(/\n(van|from|verzonden|sent|op .* schreef|on .* wrote)\s*:/i)[0];   // enkel eigen tekst, niet het citaat
    var onderwerp = ((r[0] || "") + " " + tekst).toLowerCase();
    if (REF.test(onderwerp)) return event.completed({ allowEvent: true });
    for (var c = 0; c < CODES.length; c++) if (onderwerp.indexOf(CODES[c]) >= 0) return event.completed({ allowEvent: true });
    var extern = [].concat(r[1] || [], r[2] || [], r[3] || []).filter(function (x) { return nodigNummer(x.emailAddress); });
    if (!extern.length) return event.completed({ allowEvent: true });
    event.completed({
      allowEvent: false,
      errorMessage: "Geen #nummer gevonden. Mail naar buiten = #nummer van de Trello-kaart (bv. #1159) of #info " +
        "(geen dossier). Klik 'Niet verzenden' en voeg het toe, of 'Toch verzenden'."
    });
  }, function () { event.completed({ allowEvent: true }); });   // bij een fout nooit blokkeren
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
