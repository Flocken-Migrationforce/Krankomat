

window.Krankomat = window.Krankomat || {};

Krankomat.Preview = {
    // Subject changed: removed student ID, as per request
    emailSubjectTemplate: "{art} {profilName} {Datum} [{Vornamen} {Nachname}]",

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

        // Check for Exam
        const isExam = data.absenceReasons && data.absenceReasons.exam;
        const matriculationLine = isExam && data.userData.studentId ? `Matrikelnummer: ${data.userData.studentId}\n` : '';

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

        // Templates updated to match user request exactly
        const bodyTemplate = `{anrede}

hiermit melde ich mich{wiederGesundAb} {DateText}.

Studiengruppe / Jahrgang: {profilName}
Name: {Nachname}
Vorname: {Vornamen}
{optionalMatrikelnummer}Krankmeldung: {krankX}
Gesundmeldung: {gesundX}
Erster Krankheitstag: {Datum}
Letzter Krankheitstag: {Datum2}
Attest vorhanden ja/nein: {attest}
Attest über eAU: {eau}
(freiwillig gesetzlich Versicherte RSA/RIA, AzVA, Soz. Arbeit, E-Government)
Prüfungstag/ Klausur/ Leistungsnachweis ja/nein: {prüfungstag}
Unfall (auch privat) ja/nein: {unfall}
Bemerkung / voraussichtliche Dauer: {bemerkung}

Mit freundlichen Grüßen

{Vornamen} {Nachname}`;

        const config = data.config || {};
        const profileName = config.profileName || '';

        const bemerkungRaw = data.details.comments ? data.details.comments.trim() : '';

        const context = {
            Vornamen: data.userData.firstName || '', 
            Nachname: data.userData.lastName || '', 
            profilName: profileName,
            optionalMatrikelnummer: matriculationLine,
            krankX: noticeType === 'Krankmeldung' ? 'x' : '',
            gesundX: noticeType === 'Gesundmeldung' ? 'x' : '',
            Datum: data.sicknessStartDate,
            Datum2: data.sicknessEndDate || '',
            attest: (data.absenceReasons && data.absenceReasons.attest) ? 'ja' : 'nein',
            eau: (data.absenceReasons && data.absenceReasons.eau) ? 'ja' : 'nein',
            prüfungstag: (data.absenceReasons && data.absenceReasons.exam) ? 'ja' : 'nein',
            unfall: (data.absenceReasons && data.absenceReasons.unfall) ? 'ja' : 'nein',
            bemerkung: bemerkungRaw,
            // Construct sentence parts for the intro line
            wiederGesundAb: noticeType === 'Gesundmeldung' ? ' wieder gesund ab' : ' krank für heute, den',
            DateText: noticeType === 'Gesundmeldung' ? (data.sicknessEndDate || '') : Krankomat.Utils.todayFormatted(),
            art: noticeType,
            anrede: anredeText
        };

        const body = Krankomat.Utils.renderTemplate(bodyTemplate.trim(), context).replace(/\n{3,}/g, '\n\n');
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
        const androidBtn = document.getElementById('btn-mail-android');
        const iosBtn = document.getElementById('btn-mail-ios');
        const outlookTooltip = document.getElementById('outlook-tooltip');

        if (outlookBtn) {
            if (email.toList.length > 0) {
                // Outlook Web Link (Universal Link for iOS App often intercepts this too)
                // This satisfies the "open outlook through a website or the Outlook App" requirement
                const outlookHref = `https://outlook.office.com/mail/deeplink/compose?to=${encodeURIComponent(email.to)}&subject=${encodeURIComponent(email.subject)}&body=${encodeURIComponent(email.body)}`;
                const mailtoHref = `mailto:${toMailtoValue}?subject=${encodeURIComponent(email.subject)}&body=${encodeURIComponent(email.body)}`;
                
                outlookBtn.disabled = false;
                outlookBtn.onclick = () => window.open(outlookHref, '_blank', 'noopener,noreferrer');
                if (outlookTooltip) outlookTooltip.classList.add('hidden');
                
                // Android Button: Standard mailto
                if (androidBtn) {
                    androidBtn.href = mailtoHref;
                    androidBtn.classList.remove('opacity-50', 'cursor-not-allowed');
                    androidBtn.classList.add('hover:bg-slate-50', 'dark:hover:bg-slate-600');
                }
                
                // iOS Button: Outlook Web/App Link
                if (iosBtn) {
                    iosBtn.href = "#"; 
                    iosBtn.onclick = (e) => {
                        e.preventDefault();
                        // Open Outlook Deep Link in new tab (often triggers app on mobile or goes to OWA)
                        window.open(outlookHref, '_blank', 'noopener,noreferrer');
                    };
                    iosBtn.classList.remove('opacity-50', 'cursor-not-allowed');
                    iosBtn.classList.add('hover:bg-slate-50', 'dark:hover:bg-slate-600');
                }
            } else {
                outlookBtn.disabled = true;
                outlookBtn.onclick = null;
                if (outlookTooltip) outlookTooltip.classList.remove('hidden'); 
                
                if (androidBtn) {
                    androidBtn.removeAttribute('href');
                    androidBtn.classList.add('opacity-50', 'cursor-not-allowed');
                }
                if (iosBtn) {
                    iosBtn.removeAttribute('href');
                    iosBtn.onclick = null;
                    iosBtn.classList.add('opacity-50', 'cursor-not-allowed');
                }
            }
        }
    }
};