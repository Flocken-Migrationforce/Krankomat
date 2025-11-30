
window.Krankomat = window.Krankomat || {};

Krankomat.Preview = {
    emailSubjectTemplate: "{art} E-Gov 2025 {Datum} [{Vornamen} {Nachname}, {Matrikelnummer}]",

    init: function() {
        this.setupCopyButtons();
    },

    setupCopyButtons: function() {
        const attach = (btnId, sourceId) => {
            const btn = document.getElementById(btnId);
            if (!btn) return;
            
            btn.innerHTML = Krankomat.Utils.Icons.copy;
            btn.addEventListener('click', () => {
                const el = document.getElementById(sourceId);
                const text = sourceId === 'preview-body' 
                    ? el.value 
                    : el.innerText;
                
                Krankomat.Utils.copyToClipboard(text, () => {
                    btn.innerHTML = Krankomat.Utils.Icons.check;
                    setTimeout(() => btn.innerHTML = Krankomat.Utils.Icons.copy, 2000);
                });
            });
        };
        attach('copy-subject-btn', 'preview-subject');
        attach('copy-to-btn', 'preview-to');
        attach('copy-body-btn', 'preview-body');
    },

    generate: function() {
        const data = Krankomat.State.data;
        if (!data.userData) return { subject: '', body: '', to: '', toList: [] };

        const selectedRecipients = data.recipients ? data.recipients.filter(r => r.isSelected) : [];
        const noticeType = data.sicknessEndDate ? 'Gesundmeldung' : 'Krankmeldung';
        
        const anreden = [...new Set(selectedRecipients.map(r => r.anrede.trim()))];
        const anredeText = selectedRecipients.length === 0 ? 'Sehr geehrte Damen und Herren,' : `${anreden.join(',\n')},`;

        let abwesenheitsgrund = '';
        if (noticeType === 'Krankmeldung') {
            if (data.absenceReasons.partialDay) {
                abwesenheitsgrund = 'Hiermit melde ich mich für den Rest des heutigen Tages krank.';
            } else {
                const verpasst = [];
                if (data.absenceReasons.lecture) verpasst.push('der Vorlesungszeit');
                if (data.absenceReasons.internship) verpasst.push('der Berufspraxis');
                if (data.absenceReasons.exam) verpasst.push('einer Prüfungsleistung');

                if (verpasst.length > 0) {
                    abwesenheitsgrund = `Aufgrund meiner Krankheit kann ich heute nicht an ${verpasst.join(' sowie ')} teilnehmen.`;
                }
            }
        }

        const bodyTemplate = noticeType === 'Gesundmeldung' ? 
`{anrede}

hiermit melde ich mich ab dem {Datum2} wieder gesund.
Ich war vom {Datum} bis einschließlich {Datum2} krank.

Name: {Vornamen} {Nachname}
Matrikelnummer: {Matrikelnummer}
Studiengang: {studiengang}

{bemerkung}

Mit freundlichen Grüßen
{Vornamen} {Nachname}` : 
`{anrede}

hiermit melde ich mich für den {Datum} krank.
{abwesenheitsgrund}
{dauermeldung}

Name: {Vornamen} {Nachname}
Matrikelnummer: {Matrikelnummer}
Studiengang: {studiengang}

{bemerkung}

Mit freundlichen Grüßen
{Vornamen} {Nachname}`;

        const context = {
            Vornamen: data.userData.firstName, 
            Nachname: data.userData.lastName, 
            Matrikelnummer: data.userData.studentId,
            Datum: data.sicknessStartDate, 
            Datum2: data.sicknessEndDate, 
            art: noticeType, 
            studiengang: data.userData.studyProgram || 'E-Government',
            anrede: anredeText, 
            bemerkung: data.details.comments,
            prüfungstag: data.absenceReasons.exam ? 'ja' : 'nein',
            dauermeldung: data.sicknessEndDate ? `Ich bin voraussichtlich bis einschließlich ${data.sicknessEndDate} krank.` : '',
            abwesenheitsgrund: abwesenheitsgrund,
        };

        const body = Krankomat.Utils.renderTemplate(bodyTemplate.trim(), context);
        const subject = Krankomat.Utils.renderTemplate(this.emailSubjectTemplate, context);
        const toList = selectedRecipients.map(r => r.email).filter(Boolean).filter(e => e.trim() !== '');
        const toValue = toList.join('; ');

        return { subject, body: body.trim(), to: toValue, toList };
    },

    render: function() {
        const email = this.generate();
        
        // Update UI
        const subjectEl = document.getElementById('preview-subject');
        if (subjectEl) subjectEl.innerText = email.subject;
        
        const toEl = document.getElementById('preview-to');
        if (toEl) toEl.innerText = email.to || 'Empfänger auswählen...';
        
        const bodyEl = document.getElementById('preview-body');
        if (bodyEl) bodyEl.value = email.body;

        // Buttons
        const toMailtoValue = email.toList.join(',');
        const outlookBtn = document.getElementById('open-outlook-btn');
        const mailtoLink = document.getElementById('open-mailto-link');
        const outlookTooltip = document.getElementById('outlook-tooltip');

        if (outlookBtn) {
            if (email.toList.length > 0) {
                const outlookHref = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(email.to)}&subject=${encodeURIComponent(email.subject)}&body=${encodeURIComponent(email.body)}`;
                const mailtoHref = `mailto:${toMailtoValue}?subject=${encodeURIComponent(email.subject)}&body=${encodeURIComponent(email.body)}`;
                
                outlookBtn.disabled = false;
                outlookBtn.onclick = () => window.open(outlookHref, '_blank', 'noopener,noreferrer');
                if (outlookTooltip) outlookTooltip.classList.add('hidden');
                
                if (mailtoLink) {
                    mailtoLink.href = mailtoHref;
                    mailtoLink.classList.remove('opacity-50', 'cursor-not-allowed');
                    mailtoLink.classList.add('hover:bg-slate-50', 'dark:hover:bg-slate-600');
                }
            } else {
                outlookBtn.disabled = true;
                outlookBtn.onclick = null;
                if (outlookTooltip) outlookTooltip.classList.remove('hidden'); 
                
                if (mailtoLink) {
                    mailtoLink.removeAttribute('href');
                    mailtoLink.classList.add('opacity-50', 'cursor-not-allowed');
                }
            }
        }
    }
};