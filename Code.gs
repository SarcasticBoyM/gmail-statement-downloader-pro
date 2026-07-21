// Accounts Reco V19 - Cheque-number columns forced to plain text; no rupee formatting and leading zeros preserved
// A re-presented cheque is linked as one lifecycle: earlier bounced presentation(s) remain in the bounce report,
// while the final successful credit is used for reconciliation and is not left in unmatched bank entries.
// Rules retained: normal payments +/-3 days, cheques same date then +/-10 days,
// GPay individual matching on same day first, then next day; grouped only for remaining entries.
// Retains party/narration similarity, leading-zero-safe cheque matching, and excludes bounced credits from successful matching.
const RECO_CONFIG = {
  SHEET_1213: 'Balaji Traders Axis Bank-1213',
  SHEET_2224: 'Balaji Traders Axis Bank-2224',
  SHEET_VYAPAR: 'VYAPAR_PAYMENT_IN',
  REPORT_SHEET: 'RECONCILIATION_REPORT',
  UNMATCHED_BANK_SHEET: 'UNMATCHED_BANK_ENTRIES',
  PARTY_SUMMARY_SHEET: 'PARTY_LEDGER_SUMMARY',
  SELF_ACCOUNT_SHEET: 'SELF_ACCOUNT_ENTRIES',
  BOUNCED_CHEQUE_SHEET: 'CHEQUE_BOUNCED_ENTRIES',
  VYAPAR_MISSING_EXPORT_SHEET: 'VYAPAR_MISSING_ENTRIES',

  HEADER_SCAN_ROWS: 100,
  AMOUNT_TOLERANCE: 0.50,
  DIRECT_DATE_TOLERANCE_DAYS: 3,
  CHEQUE_PRIMARY_DAYS: 10,
  CHEQUE_FALLBACK_DAYS: 30,
  GPAY_MAX_DAY_OFFSET: 1,

  // Party/narration similarity is used only as a tie-breaker when more than one
  // exact-amount candidate exists on the nearest eligible date. A unique
  // exact-amount candidate is still matched even when the narration score is low.
  PARTY_SIMILARITY_MIN_AUTO_MATCH: 20,
  PARTY_SIMILARITY_MIN_LEAD: 10,

  // Conservative thresholds used only to suggest a party name for an unmatched
  // bank credit in the export-ready Vyapar missing entries sheet.
  MISSING_PARTY_SIMILARITY_MIN: 45,
  MISSING_PARTY_SIMILARITY_MIN_LEAD: 10
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RECO')
    .addItem('Run Reconciliation', 'runReconciliation')
    .addItem('Clear Reco Reports', 'clearRecoReports')
    .addToUi();
}

function runReconciliation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Kolkata';

  const vyaparSheet = getRequiredSheet_(ss, RECO_CONFIG.SHEET_VYAPAR);
  const bank1213Sheet = getRequiredSheet_(ss, RECO_CONFIG.SHEET_1213);
  const bank2224Sheet = getRequiredSheet_(ss, RECO_CONFIG.SHEET_2224);

  const vyapar = readVyapar_(vyaparSheet, tz);
  const bank1213 = readBank_(bank1213Sheet, '1213', tz);
  const bank2224 = readBank_(bank2224Sheet, '2224', tz);
  const allBankTransactions = [...bank1213.transactions, ...bank2224.transactions];
  const allBankDebits = [...bank1213.debitTransactions, ...bank2224.debitTransactions];
  const results = new Map();

  // 0) Reserve cheque lifecycles first when the same cheque was deposited,
  // bounced, re-deposited and later cleared. This is more specific than amount-only
  // matching and prevents the final successful credit from remaining unmatched or
  // being consumed by another payment rule.
  processRepresentedChequeEntries_(
    vyapar.entries.filter(entry => entry.rule === 'CHEQUE'),
    allBankTransactions,
    results,
    tz
  );

  // 1) GPay: reserve same-day / next-day credits after represented cheques.
  processGpayEntries_(
    vyapar.entries.filter(entry => entry.rule === 'GPAY'),
    bank1213.transactions,
    results,
    tz
  );

  // 2) Process all remaining cheques using the normal cheque rules.
  processChequeEntries_(
    vyapar.entries.filter(entry => entry.rule === 'CHEQUE'),
    allBankTransactions,
    results,
    tz
  );

  // 3) Other bank payments: exact amount in the selected account within +/- 3 days.
  //    The nearest transaction date is preferred.
  processDirectEntries_(
    vyapar.entries.filter(entry => entry.rule === 'DIRECT_1213' || entry.rule === 'DIRECT_2224'),
    bank1213,
    bank2224,
    results,
    tz
  );

  // 4) Unsupported or blank payment type.
  vyapar.entries
    .filter(entry => entry.rule === 'UNKNOWN')
    .forEach(entry => {
      results.set(entry.rowNumber, {
        entry,
        ruleText: 'No rule configured for this Payment Type',
        groupTotal: '',
        bankAccount: '',
        bankTxn: null,
        difference: '',
        status: 'MANUAL CHECK',
        remarks: `Payment Type not recognised: ${entry.paymentType || '(blank)'}`
      });
    });

  writeRecoReport_(ss, vyapar.entries, results);
  writeUnmatchedBankReport_(ss, allBankTransactions);
  writeSelfAccountReport_(ss, allBankTransactions);
  writeBouncedChequeReport_(ss, allBankTransactions);
  writeVyaparMissingEntriesExport_(ss, allBankTransactions, allBankDebits, vyapar.entries);
  writePartyLedgerSummary_(ss, vyapar.entries, results);

  // Use a spreadsheet toast instead of getUi().alert().
  // getUi() is unavailable in some execution contexts, including certain editor/trigger runs.
  try {
    ss.toast(
      'Reconciliation, unmatched bank, self-account, bounced-cheque, Payment In/Out Vyapar-missing export and party-summary reports have been created/updated.',
      'Reconciliation completed',
      8
    );
  } catch (error) {
    console.log('Reconciliation completed. Reports were created/updated.');
  }
}

function clearRecoReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [RECO_CONFIG.REPORT_SHEET, RECO_CONFIG.UNMATCHED_BANK_SHEET, RECO_CONFIG.SELF_ACCOUNT_SHEET, RECO_CONFIG.BOUNCED_CHEQUE_SHEET, RECO_CONFIG.VYAPAR_MISSING_EXPORT_SHEET, RECO_CONFIG.PARTY_SUMMARY_SHEET].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
}

function processDirectEntries_(directEntries, bank1213, bank2224, results, tz) {
  directEntries
    .slice()
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(entry => {
      const bank = entry.rule === 'DIRECT_1213' ? bank1213 : bank2224;
      const allCandidates = findCandidates_(
        bank.transactions,
        entry.amount,
        entry.date,
        -RECO_CONFIG.DIRECT_DATE_TOLERANCE_DAYS,
        RECO_CONFIG.DIRECT_DATE_TOLERANCE_DAYS,
        tz
      );

      // Same date is preferred, then +/-1, +/-2 and +/-3 days. When the nearest
      // eligible date has duplicate exact amounts, Party Name vs Bank Narration
      // similarity is used as the tie-breaker.
      const nearestCandidates = selectNearestDateCandidates_(allCandidates, entry.date);
      const resolution = buildPartyAwareSingleResult_(
        entry,
        nearestCandidates,
        bank.accountName,
        `Selected account: exact amount searched within +/-${RECO_CONFIG.DIRECT_DATE_TOLERANCE_DAYS} days; nearest date preferred; narration used for duplicate amounts`
      );
      results.set(entry.rowNumber, resolution);
      if (resolution.status === 'MATCHED' && resolution.bankTxn) {
        resolution.bankTxn.used = true;
        resolution.bankTxn.usedBy = String(entry.rowNumber);
      }
    });
}

function processGpayEntries_(gpayEntries, bankTransactions, results, tz) {
  const groups = groupEntriesByDate_(gpayEntries, tz);

  Object.keys(groups).sort().forEach(dateKey => {
    const dayEntries = groups[dateKey].slice().sort((a, b) => a.rowNumber - b.rowNumber);
    const vyaparDate = dayEntries[0].date;

    // First match every GPay entry individually using an exact amount.
    // Search order is: same day first, then next day.
    const entriesByAmount = dayEntries.reduce((map, entry) => {
      const key = amountKey_(entry.amount);
      if (!map[key]) map[key] = [];
      map[key].push(entry);
      return map;
    }, {});

    Object.keys(entriesByAmount).forEach(amountKey => {
      const sameAmountEntries = entriesByAmount[amountKey]
        .slice()
        .sort((a, b) => a.rowNumber - b.rowNumber);
      const amount = sameAmountEntries[0].amount;

      // Same-day credits are allocated first. Next-day credits are allocated only
      // after same-day candidates are exhausted. Within each day, Party Name vs
      // Bank Narration similarity is used to pair duplicate equal amounts.
      for (let dayOffset = 0; dayOffset <= RECO_CONFIG.GPAY_MAX_DAY_OFFSET; dayOffset++) {
        const remainingEntries = sameAmountEntries.filter(entry => !results.has(entry.rowNumber));
        if (!remainingEntries.length) break;

        const dayCandidates = findCandidates_(
          bankTransactions,
          amount,
          vyaparDate,
          dayOffset,
          dayOffset,
          tz
        ).sort((a, b) => a.rowNumber - b.rowNumber);

        const pairs = pairEntriesAndTransactionsByPartySimilarity_(remainingEntries, dayCandidates);
        pairs.forEach(pair => {
          const entry = pair.entry;
          const txn = pair.txn;
          txn.used = true;
          txn.usedBy = String(entry.rowNumber);

          results.set(entry.rowNumber, {
            entry,
            ruleText: 'GPay: individual exact amount searched on same day first, then next day in 1213; narration used to pair duplicate equal amounts',
            groupTotal: '',
            bankAccount: txn.accountName,
            bankTxn: txn,
            difference: round2_(txn.amount - entry.amount),
            status: 'MATCHED',
            remarks: dayOffset === 0
              ? 'Same-day individual GPay credit matched before grouped matching.'
              : 'Next-day individual GPay credit matched before grouped matching.'
          });
        });
      }
    });

    // Group only entries for which no individual same-day or next-day credit remains.
    const remaining = dayEntries.filter(entry => !results.has(entry.rowNumber));
    if (!remaining.length) return;

    // A single remaining entry is not treated as a group.
    if (remaining.length === 1) {
      const entry = remaining[0];
      results.set(entry.rowNumber, {
        entry,
        ruleText: 'GPay: individual exact amount searched on same day and next day in 1213; grouping not applicable to one entry',
        groupTotal: '',
        bankAccount: '1213',
        bankTxn: null,
        difference: '',
        status: 'UNMATCHED',
        remarks: `No unused same-day or next-day 1213 credit of ₹${entry.amount.toFixed(2)} was found.`
      });
      return;
    }

    const groupTotal = round2_(remaining.reduce((sum, entry) => sum + entry.amount, 0));
    const allGroupCandidates = findCandidates_(
      bankTransactions,
      groupTotal,
      vyaparDate,
      0,
      RECO_CONFIG.GPAY_MAX_DAY_OFFSET,
      tz
    );

    // Same-day grouped credit gets priority. Next-day is considered only when
    // no same-day grouped credit exists.
    const preferredGroupCandidates = selectPreferredForwardDateCandidates_(
      allGroupCandidates,
      vyaparDate
    );

    let status = 'UNMATCHED';
    let remarks = `No same-day or next-day 1213 grouped credit of ₹${groupTotal.toFixed(2)} found after individual matching.`;
    let bankTxn = null;

    if (preferredGroupCandidates.length === 1) {
      bankTxn = preferredGroupCandidates[0];
      const matchedDayOffset = daysBetween_(vyaparDate, bankTxn.date);
      bankTxn.used = true;
      bankTxn.usedBy = remaining.map(entry => entry.rowNumber).join(',');
      status = 'MATCHED';
      remarks = `${remaining.length} GPay entries were grouped only after all available individual same-day/next-day credits were matched. Group matched on ${matchedDayOffset === 0 ? 'same day' : 'next day'}.`;
    } else if (preferredGroupCandidates.length > 1) {
      const matchedDayOffset = daysBetween_(vyaparDate, preferredGroupCandidates[0].date);
      status = 'MANUAL CHECK';
      remarks = `${preferredGroupCandidates.length} ${matchedDayOffset === 0 ? 'same-day' : 'next-day'} 1213 credits of grouped amount ₹${groupTotal.toFixed(2)} found. Bank rows: ${preferredGroupCandidates.map(txn => txn.rowNumber).join(', ')}`;
    }

    remaining.forEach(entry => {
      results.set(entry.rowNumber, {
        entry,
        ruleText: 'GPay: individual same-day first, then next-day; only remaining entries grouped using same-day first, then next-day in 1213',
        groupTotal,
        bankAccount: bankTxn ? bankTxn.accountName : '1213',
        bankTxn,
        difference: bankTxn ? round2_(bankTxn.amount - groupTotal) : '',
        status,
        remarks
      });
    });
  });
}

function processRepresentedChequeEntries_(chequeEntries, allBankTransactions, results, tz) {
  chequeEntries
    .filter(entry => entry.chequeTokens && entry.chequeTokens.length)
    .slice()
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(entry => {
      const allowedDates = buildAllowedDateSet_(
        entry.date,
        -RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
        RECO_CONFIG.CHEQUE_FALLBACK_DAYS,
        tz
      );

      // Every credit presentation of this exact cheque and amount, including
      // bounced presentations. Leading zeroes in cheque numbers are ignored.
      const presentations = allBankTransactions
        .filter(txn =>
          !txn.isSelfAccount &&
          allowedDates.has(txn.dateKey) &&
          Math.abs(txn.amount - entry.amount) <= RECO_CONFIG.AMOUNT_TOLERANCE &&
          transactionContainsChequeToken_(txn, entry.chequeTokens)
        )
        .sort((a, b) => a.date - b.date || a.accountName.localeCompare(b.accountName) || a.rowNumber - b.rowNumber);

      const bounced = presentations.filter(txn => txn.isChequeBounced);
      if (!bounced.length) return;

      // Use a successful credit only when it occurred after the final bounced
      // presentation. This represents the re-deposit that ultimately cleared.
      const lastBouncedPresentation = bounced
        .slice()
        .sort((a, b) => b.date - a.date || b.rowNumber - a.rowNumber)[0];

      const successfulAfterLastBounce = presentations.filter(txn =>
        !txn.isChequeBounced &&
        !txn.used &&
        (txn.date > lastBouncedPresentation.date ||
          (txn.date.getTime() === lastBouncedPresentation.date.getTime() && txn.rowNumber > lastBouncedPresentation.rowNumber))
      );

      if (!successfulAfterLastBounce.length) return;

      // The first successful presentation after the last bounce is the final
      // clearance. If that date contains duplicates, use party/narration
      // similarity; otherwise keep it for manual review rather than guessing.
      const firstSuccessDate = successfulAfterLastBounce[0].date;
      const firstSuccessCandidates = successfulAfterLastBounce.filter(txn =>
        txn.date.getTime() === firstSuccessDate.getTime()
      );

      let selected = null;
      let selectionNote = '';

      if (firstSuccessCandidates.length === 1) {
        selected = firstSuccessCandidates[0];
        selectionNote = 'One successful re-deposit was found after the last bounced presentation.';
      } else {
        const choice = chooseCandidateByPartySimilarity_(entry, firstSuccessCandidates);
        if (choice.autoMatch) {
          selected = choice.txn;
          selectionNote = `${firstSuccessCandidates.length} successful re-deposits existed on the clearance date; party/narration similarity selected bank row ${selected.rowNumber} (${choice.score}%).`;
        } else {
          results.set(entry.rowNumber, {
            entry,
            ruleText: 'Cheque lifecycle: deposit -> bounce -> re-deposit -> clearance; duplicate successful re-deposits require review',
            groupTotal: '',
            bankAccount: '1213 / 2224',
            bankTxn: null,
            difference: '',
            status: 'MANUAL CHECK',
            remarks: `${firstSuccessCandidates.length} unused successful credits were found after the last bounce for cheque no ${entry.chequeNo} and amount ₹${entry.amount.toFixed(2)}. Rows: ${firstSuccessCandidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
          });
          return;
        }
      }

      selected.used = true;
      selected.usedBy = String(entry.rowNumber);

      const relatedBounces = bounced.filter(txn =>
        txn.date < selected.date ||
        (txn.date.getTime() === selected.date.getTime() && txn.rowNumber < selected.rowNumber)
      );
      relatedBounces.forEach(txn => {
        txn.usedBy = appendUsage_(txn.usedBy, entry.rowNumber);
      });

      const bounceDetails = relatedBounces.map(txn => {
        const creditDate = Utilities.formatDate(txn.date, tz || 'Asia/Kolkata', 'dd-MMM-yyyy');
        const reversalDate = txn.bounceDate
          ? Utilities.formatDate(txn.bounceDate, tz || 'Asia/Kolkata', 'dd-MMM-yyyy')
          : creditDate;
        return `${txn.accountName} credit row ${txn.rowNumber} on ${creditDate}, reversed at row ${txn.bounceDebitRow || ''} on ${reversalDate}`;
      }).join('; ');
      const clearanceDate = Utilities.formatDate(selected.date, tz || 'Asia/Kolkata', 'dd-MMM-yyyy');

      results.set(entry.rowNumber, buildChequeResult_(
        entry,
        selected,
        'Cheque lifecycle: earlier bounced presentation(s) linked; final successful re-deposit matched by cheque no + exact amount',
        `${selectionNote} Earlier presentation(s): ${bounceDetails}. Final clearance: ${selected.accountName} row ${selected.rowNumber} on ${clearanceDate}.`
      ));
    });
}

function processChequeEntries_(chequeEntries, allBankTransactions, results, tz) {
  // A credit that was reversed as a cheque bounce must not block a later,
  // successful re-presentation of the same cheque. Therefore successful and
  // bounced credits are searched separately. Example: cheque 000692 credited
  // and returned on 03-Apr, then finally cleared on 15-Apr: the 15-Apr credit
  // is matched, while the 03-Apr pair remains visible in CHEQUE_BOUNCED_ENTRIES.
  const successfulTransactions = allBankTransactions.filter(txn => !txn.isChequeBounced);
  const bouncedTransactions = allBankTransactions.filter(txn => txn.isChequeBounced);

  chequeEntries
    .filter(entry => !results.has(entry.rowNumber))
    .slice()
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(entry => {
      // Stage 1: same-date exact amount, but only against credits that were not reversed.
      const sameDayAmountCandidates = findCandidates_(
        successfulTransactions,
        entry.amount,
        entry.date,
        0,
        0,
        tz
      ).sort((a, b) => a.accountName.localeCompare(b.accountName) || a.rowNumber - b.rowNumber);

      if (sameDayAmountCandidates.length === 1) {
        const txn = sameDayAmountCandidates[0];
        txn.used = true;
        txn.usedBy = String(entry.rowNumber);
        const bounceNote = annotateEarlierBouncedCheque_(entry, txn, bouncedTransactions, tz);
        results.set(entry.rowNumber, buildChequeResult_(
          entry,
          txn,
          'Cheque: unique same-date successful credit matched; cheque no not required',
          `Only one unused same-date successful credit had this amount.${bounceNote}`
        ));
        return;
      }

      if (sameDayAmountCandidates.length > 1) {
        resolveChequeDuplicates_(
          entry,
          sameDayAmountCandidates,
          results,
          'Cheque: duplicate same-date successful credits resolved using cheque no + amount'
        );
        return;
      }

      // Stage 2: no successful same-date credit. Search exact amount from 10 days
      // before through 10 days after. Reversed/bounced credits are excluded.
      const primaryWindowCandidates = findCandidates_(
        successfulTransactions,
        entry.amount,
        entry.date,
        -RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
        RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
        tz
      );
      const nearestPrimaryCandidates = selectNearestDateCandidates_(primaryWindowCandidates, entry.date)
        .sort((a, b) => a.accountName.localeCompare(b.accountName) || a.rowNumber - b.rowNumber);

      if (nearestPrimaryCandidates.length === 1) {
        const txn = nearestPrimaryCandidates[0];
        txn.used = true;
        txn.usedBy = String(entry.rowNumber);
        const dayDifference = daysBetween_(entry.date, txn.date);
        const dateDirection = dayDifference < 0
          ? `${Math.abs(dayDifference)} day(s) before`
          : `${dayDifference} day(s) after`;
        const bounceNote = annotateEarlierBouncedCheque_(entry, txn, bouncedTransactions, tz);
        results.set(entry.rowNumber, buildChequeResult_(
          entry,
          txn,
          `Cheque: unique successful exact amount matched within +/-${RECO_CONFIG.CHEQUE_PRIMARY_DAYS} days; nearest date preferred`,
          `No same-date successful credit was found; one unique successful amount match was found ${dateDirection} the Vyapar date.${bounceNote}`
        ));
        return;
      }

      if (nearestPrimaryCandidates.length > 1) {
        resolveChequeDuplicates_(
          entry,
          nearestPrimaryCandidates,
          results,
          `Cheque: duplicate successful exact amounts on nearest date within +/-${RECO_CONFIG.CHEQUE_PRIMARY_DAYS} days resolved using cheque no + amount`
        );
        return;
      }

      // Stage 3: only when no successful amount match exists within +/-10 days,
      // search day 11-30 after the Vyapar date using normalized cheque no + amount.
      if (entry.chequeTokens.length) {
        const fallbackCandidates = findChequeCandidates_(
          successfulTransactions,
          entry,
          RECO_CONFIG.CHEQUE_PRIMARY_DAYS + 1,
          RECO_CONFIG.CHEQUE_FALLBACK_DAYS,
          tz
        );

        if (fallbackCandidates.length === 1) {
          const txn = fallbackCandidates[0];
          txn.used = true;
          txn.usedBy = String(entry.rowNumber);
          const bounceNote = annotateEarlierBouncedCheque_(entry, txn, bouncedTransactions, tz);
          results.set(entry.rowNumber, buildChequeResult_(
            entry,
            txn,
            `Cheque fallback: successful cheque no + amount matched from day ${RECO_CONFIG.CHEQUE_PRIMARY_DAYS + 1} to ${RECO_CONFIG.CHEQUE_FALLBACK_DAYS}`,
            `No successful amount match was found within +/-${RECO_CONFIG.CHEQUE_PRIMARY_DAYS} days; normalized cheque no ${entry.chequeNo} and amount matched during the forward extended search (leading zeroes ignored).${bounceNote}`
          ));
          return;
        }

        if (fallbackCandidates.length > 1) {
          const similarityChoice = chooseCandidateByPartySimilarity_(entry, fallbackCandidates);
          if (similarityChoice.autoMatch) {
            const txn = similarityChoice.txn;
            txn.used = true;
            txn.usedBy = String(entry.rowNumber);
            const bounceNote = annotateEarlierBouncedCheque_(entry, txn, bouncedTransactions, tz);
            results.set(entry.rowNumber, buildChequeResult_(
              entry,
              txn,
              'Cheque fallback: duplicate successful cheque no + amount matches resolved using party/narration similarity',
              `${fallbackCandidates.length} successful credits matched cheque no ${entry.chequeNo} and amount; bank row ${txn.rowNumber} was selected with ${similarityChoice.score}% party/narration similarity.${bounceNote}`
            ));
            return;
          }

          results.set(entry.rowNumber, buildChequeManualResult_(
            entry,
            fallbackCandidates,
            `Cheque fallback: duplicate successful cheque no + amount matches from day ${RECO_CONFIG.CHEQUE_PRIMARY_DAYS + 1} to ${RECO_CONFIG.CHEQUE_FALLBACK_DAYS}`
          ));
          return;
        }
      }

      // No successful clearance was found. Only now check whether the cheque was
      // credited and reversed. This reports CHEQUE BOUNCED instead of UNMATCHED.
      const bouncedCandidates = entry.chequeTokens.length
        ? findChequeCandidates_(
            bouncedTransactions,
            entry,
            -RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
            RECO_CONFIG.CHEQUE_FALLBACK_DAYS,
            tz
          )
        : findCandidates_(
            bouncedTransactions,
            entry.amount,
            entry.date,
            -RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
            RECO_CONFIG.CHEQUE_PRIMARY_DAYS,
            tz,
            { includeBounced: true }
          );

      if (bouncedCandidates.length) {
        const nearestBounced = selectNearestDateCandidates_(bouncedCandidates, entry.date);
        const txn = nearestBounced[0];
        txn.usedBy = appendUsage_(txn.usedBy, entry.rowNumber);
        results.set(entry.rowNumber, buildChequeResult_(
          entry,
          txn,
          'Cheque: no successful clearance found; reversed bank credit detected',
          'The cheque credit was reversed and no later successful clearance matching the configured rules was found.'
        ));
        return;
      }

      results.set(entry.rowNumber, {
        entry,
        ruleText: `Cheque: successful same-date amount, then successful amount within +/-${RECO_CONFIG.CHEQUE_PRIMARY_DAYS} days, then successful cheque no + amount from day ${RECO_CONFIG.CHEQUE_PRIMARY_DAYS + 1} to ${RECO_CONFIG.CHEQUE_FALLBACK_DAYS} after the Vyapar date`,
        groupTotal: '',
        bankAccount: '1213 / 2224',
        bankTxn: null,
        difference: '',
        status: entry.chequeTokens.length ? 'UNMATCHED' : 'MANUAL CHECK',
        remarks: entry.chequeTokens.length
          ? `No unused successful bank credit was found for ₹${entry.amount.toFixed(2)} and cheque no ${entry.chequeNo || '(not available)'}. Bounced credits were also checked separately.`
          : 'No successful amount match was found and the cheque number could not be extracted for the 30-day forward fallback search.'
      });
    });
}

function annotateEarlierBouncedCheque_(entry, successfulTxn, bouncedTransactions, tz) {
  if (!entry.chequeTokens.length) return '';

  const related = bouncedTransactions
    .filter(txn =>
      Math.abs(txn.amount - entry.amount) <= RECO_CONFIG.AMOUNT_TOLERANCE &&
      txn.date <= successfulTxn.date &&
      transactionContainsChequeToken_(txn, entry.chequeTokens)
    )
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber);

  if (!related.length) return '';

  related.forEach(txn => {
    txn.usedBy = appendUsage_(txn.usedBy, entry.rowNumber);
  });

  const details = related
    .map(txn => `${txn.accountName} row ${txn.rowNumber} on ${Utilities.formatDate(txn.date, tz || 'Asia/Kolkata', 'dd-MMM-yyyy')}`)
    .join(', ');
  return ` Earlier bounced presentation(s) were detected at ${details}; the later successful credit was used for reconciliation.`;
}

function appendUsage_(existing, rowNumber) {
  const parts = String(existing || '')
    .split(',')
    .map(value => value.trim())
    .filter(Boolean);
  const value = String(rowNumber);
  if (!parts.includes(value)) parts.push(value);
  return parts.join(',');
}

function resolveChequeDuplicates_(entry, amountCandidates, results, ruleText) {
  if (!entry.chequeTokens.length) {
    results.set(entry.rowNumber, {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: '1213 / 2224',
      bankTxn: null,
      difference: '',
      status: 'MANUAL CHECK',
      remarks: `${amountCandidates.length} bank credits have the same amount, but the Vyapar cheque number is not available. Rows: ${amountCandidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
    });
    return;
  }

  const chequeCandidates = amountCandidates.filter(txn =>
    transactionContainsChequeToken_(txn, entry.chequeTokens)
  );

  if (chequeCandidates.length === 1) {
    const txn = chequeCandidates[0];
    txn.used = true;
    txn.usedBy = String(entry.rowNumber);
    results.set(entry.rowNumber, buildChequeResult_(
      entry,
      txn,
      ruleText,
      `There were ${amountCandidates.length} duplicate amount credits; cheque no ${entry.chequeNo} identified one unique bank transaction.`
    ));
    return;
  }

  if (chequeCandidates.length > 1) {
    const similarityChoice = chooseCandidateByPartySimilarity_(entry, chequeCandidates);
    if (similarityChoice.autoMatch) {
      const txn = similarityChoice.txn;
      txn.used = true;
      txn.usedBy = String(entry.rowNumber);
      results.set(entry.rowNumber, buildChequeResult_(
        entry,
        txn,
        `${ruleText}; remaining duplicate cheque matches resolved using party/narration similarity`,
        `${chequeCandidates.length} credits matched cheque no ${entry.chequeNo}; party/narration similarity selected bank row ${txn.rowNumber} (${similarityChoice.score}%).`
      ));
      return;
    }
  }

  results.set(entry.rowNumber, {
    entry,
    ruleText,
    groupTotal: '',
    bankAccount: '1213 / 2224',
    bankTxn: null,
    difference: '',
    status: 'MANUAL CHECK',
    remarks: chequeCandidates.length > 1
      ? `${chequeCandidates.length} duplicate credits also match cheque no ${entry.chequeNo}, and narration did not identify one confidently. Rows: ${chequeCandidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
      : `${amountCandidates.length} duplicate amount credits were found, but none contains cheque no ${entry.chequeNo}. Amount rows: ${amountCandidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
  });
}

function buildChequeResult_(entry, txn, ruleText, remarks) {
  if (txn.isChequeBounced) {
    const bounceDetail = txn.bounceNarration || 'Debit reversal found for the same cheque credit.';
    return {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: txn.accountName,
      bankTxn: txn,
      difference: round2_(txn.amount - entry.amount),
      status: 'CHEQUE BOUNCED',
      remarks: `${remarks} CHEQUE BOUNCED: credit was reversed on bank row ${txn.bounceDebitRow || ''}. ${bounceDetail}`.trim()
    };
  }

  return {
    entry,
    ruleText,
    groupTotal: '',
    bankAccount: txn.accountName,
    bankTxn: txn,
    difference: round2_(txn.amount - entry.amount),
    status: 'MATCHED',
    remarks
  };
}

function buildChequeManualResult_(entry, candidates, ruleText) {
  return {
    entry,
    ruleText,
    groupTotal: '',
    bankAccount: '1213 / 2224',
    bankTxn: null,
    difference: '',
    status: 'MANUAL CHECK',
    remarks: `${candidates.length} credits have the same cheque no and amount. Rows: ${candidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
  };
}

function readVyapar_(sheet, tz) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error(`${sheet.getName()} is empty.`);

  const aliases = {
    date: ['date', 'payment date', 'entry date'],
    refNo: ['ref no', 'reference no', 'reference number', 'ref number', 'voucher no', 'cheque no', 'chq no'],
    party: ['party', 'party name', 'customer', 'customer name'],
    entryType: ['entry type', 'type'],
    amount: ['received', 'received amount', 'receipt amount', 'total amt', 'total amount', 'amount', 'payment amount'],
    paymentType: ['payment type', 'payment mode', 'mode of payment']
  };

  const headerRowIndex = findHeaderRow_(data, aliases, ['date', 'amount', 'paymentType']);
  const headers = data[headerRowIndex];
  const cols = mapColumns_(headers, aliases, ['date', 'refNo', 'party', 'entryType', 'amount', 'paymentType'], sheet.getName());

  const entries = [];
  for (let r = headerRowIndex + 1; r < data.length; r++) {
    const row = data[r];
    const date = parseDate_(row[cols.date]);
    const amount = parseAmount_(row[cols.amount]);
    const refNo = String(row[cols.refNo] ?? '').trim();
    const paymentType = String(row[cols.paymentType] ?? '').trim();

    if (!date || !isFinite(amount) || amount === 0) continue;

    const chequeTokens = extractVyaparChequeTokens_(refNo, paymentType);
    entries.push({
      rowNumber: r + 1,
      date,
      dateKey: dateKey_(date, tz),
      refNo,
      chequeTokens,
      chequeNo: chequeTokens.length ? chequeTokens[0] : '',
      party: String(row[cols.party] ?? '').trim(),
      entryType: String(row[cols.entryType] ?? '').trim(),
      amount: Math.abs(round2_(amount)),
      paymentType,
      rule: classifyPaymentType_(paymentType)
    });
  }

  return { headerRowIndex, entries };
}

function readBank_(sheet, accountSuffix, tz) {
  const data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error(`${sheet.getName()} is empty.`);

  const aliases = {
    date: ['transaction date', 'txn date', 'value date', 'date', 'tran date', 'posting date'],
    narration: ['narration', 'description', 'particulars', 'transaction remarks', 'remarks', 'details'],
    chequeNo: ['cheque no', 'cheque number', 'chq no', 'chq number', 'chqno', 'instrument no', 'instrument number', 'cheque/ref no', 'cheque ref no'],
    credit: ['credit', 'credit amount', 'deposit', 'deposit amount', 'cr amount', 'amount credited', 'credit inr', 'total cr'],
    debit: ['debit', 'debit amount', 'withdrawal', 'withdrawal amount', 'dr amount', 'amount debited', 'debit inr', 'total dr'],
    amount: ['amount', 'transaction amount', 'txn amount', 'amount inr'],
    type: ['type', 'transaction type', 'dr cr', 'dr/cr', 'cr/dr']
  };

  const headerRowIndex = findBankHeaderRow_(data, aliases);
  const headers = data[headerRowIndex];
  const dateCol = findColumn_(headers, aliases.date);
  const narrationCol = findColumn_(headers, aliases.narration);
  const chequeNoCol = findColumn_(headers, aliases.chequeNo);
  const creditCol = findColumn_(headers, aliases.credit);
  const debitCol = findColumn_(headers, aliases.debit);
  const amountCol = findColumn_(headers, aliases.amount);
  const typeCol = findColumn_(headers, aliases.type);

  if (dateCol === -1) {
    throw new Error(`Date column not found in ${sheet.getName()}. Headers found: ${headers.join(' | ')}`);
  }
  if (creditCol === -1 && debitCol === -1 && amountCol === -1) {
    throw new Error(`Credit/Debit/Amount column not found in ${sheet.getName()}. Headers found: ${headers.join(' | ')}`);
  }

  const transactions = [];
  const debitTransactions = [];

  for (let r = headerRowIndex + 1; r < data.length; r++) {
    const row = data[r];
    const date = parseDate_(row[dateCol]);
    if (!date) continue;

    const narration = narrationCol !== -1 ? String(row[narrationCol] ?? '').trim() : '';
    const bankChequeField = chequeNoCol !== -1 ? String(row[chequeNoCol] ?? '').trim() : '';
    const chequeTokens = extractBankChequeTokens_(bankChequeField, narration);

    let creditAmount = creditCol !== -1 ? parseAmount_(row[creditCol]) : NaN;
    let debitAmount = debitCol !== -1 ? parseAmount_(row[debitCol]) : NaN;

    if (amountCol !== -1 && creditCol === -1 && debitCol === -1) {
      const rawAmount = parseAmount_(row[amountCol]);
      const typeText = typeCol !== -1 ? normalize_(row[typeCol]) : '';
      const rawText = normalize_(row[amountCol]);
      const looksDebit = /(^|\s)(dr|debit)(\s|$)/.test(typeText) || /(^|\s)(dr|debit)(\s|$)/.test(rawText);
      const looksCredit = /(^|\s)(cr|credit)(\s|$)/.test(typeText) || /(^|\s)(cr|credit)(\s|$)/.test(rawText);

      if (looksDebit || rawAmount < 0) debitAmount = Math.abs(rawAmount);
      else if (looksCredit || rawAmount > 0) creditAmount = Math.abs(rawAmount);
    }

    if (isFinite(creditAmount) && creditAmount > 0) {
      transactions.push({
        accountName: accountSuffix,
        sheetName: sheet.getName(),
        rowNumber: r + 1,
        date,
        dateKey: dateKey_(date, tz),
        narration,
        chequeTokens,
        chequeNo: chequeTokens.length ? chequeTokens[0] : bankChequeField,
        amount: round2_(Math.abs(creditAmount)),
        used: false,
        usedBy: '',
        isSelfAccount: isSelfAccountNarration_(narration, accountSuffix),
        isChequeBounced: false,
        bounceDebitRow: '',
        bounceDate: null,
        bounceNarration: '',
        bounceAmount: ''
      });
    }

    if (isFinite(debitAmount) && debitAmount > 0) {
      debitTransactions.push({
        accountName: accountSuffix,
        sheetName: sheet.getName(),
        rowNumber: r + 1,
        date,
        dateKey: dateKey_(date, tz),
        narration,
        chequeTokens,
        chequeNo: chequeTokens.length ? chequeTokens[0] : bankChequeField,
        amount: round2_(Math.abs(debitAmount)),
        isSelfAccount: isSelfAccountNarration_(narration, accountSuffix),
        isChequeBounceReversal: false,
        bounceCreditRow: ''
      });
    }
  }

  detectChequeBounces_(transactions, debitTransactions);
  return { accountName: accountSuffix, transactions, debitTransactions };
}

function detectChequeBounces_(creditTransactions, debitTransactions) {
  debitTransactions
    .slice()
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber)
    .forEach(debit => {
      const bounceNarration = isChequeBounceNarration_(debit.narration);
      let candidates = creditTransactions.filter(credit =>
        !credit.isChequeBounced &&
        credit.dateKey === debit.dateKey &&
        credit.rowNumber < debit.rowNumber &&
        Math.abs(credit.amount - debit.amount) <= RECO_CONFIG.AMOUNT_TOLERANCE
      );

      if (!candidates.length) return;

      const tokenCandidates = debit.chequeTokens && debit.chequeTokens.length
        ? candidates.filter(credit => transactionContainsChequeToken_(credit, debit.chequeTokens))
        : [];

      if (tokenCandidates.length) candidates = tokenCandidates;
      else if (!bounceNarration) return;

      // The reversal normally follows the original credit, so use the closest prior credit row.
      candidates.sort((a, b) => b.rowNumber - a.rowNumber);
      const credit = candidates[0];
      credit.isChequeBounced = true;
      credit.bounceDebitRow = debit.rowNumber;
      credit.bounceDate = debit.date;
      credit.bounceNarration = debit.narration;
      credit.bounceAmount = debit.amount;

      // Mark the matching debit as a cheque-bounce reversal so it is not exported
      // as a genuine Payment Out transaction in VYAPAR_MISSING_ENTRIES.
      debit.isChequeBounceReversal = true;
      debit.bounceCreditRow = credit.rowNumber;
    });
}

function isChequeBounceNarration_(narration) {
  const text = String(narration || '').toUpperCase();
  return /(BOUNCE|BOUNCED|RETURN|RETURNED|RTN|REJECT|REJECTED|FUNDS?\s+INSUFFICIENT|INSUFFICIENT\s+FUNDS|CHQ\s*RET|CHEQUE\s*RET|CLG\s*RET|OUTWARD\s*RETURN)/.test(text);
}

function isSelfAccountNarration_(narration, accountSuffix) {
  const raw = String(narration || '').toUpperCase();
  const normalized = normalize_(narration);
  const compact = raw.replace(/[^A-Z0-9]/g, '');

  const explicitSelfTransfer =
    /(^|[^A-Z])SELFFT([^A-Z]|$)/.test(raw) ||
    /(^|[^A-Z])SELF\s*(FT|TRF|TRANSFER)([^A-Z]|$)/.test(raw) ||
    /\bOWN\s*(ACCOUNT|A\/?C)\b/.test(raw) ||
    /\bINTERNAL\s+TRANSFER\b/.test(raw) ||
    /\bTRANSFER\s+BETWEEN\s+OWN\s+ACCOUNTS?\b/.test(raw) ||
    normalized.includes('self account') ||
    compact.includes('SELFTRANSFER');

  // Axis mobile self-transfer narrations often contain MOB/SELFFT and the other own-account suffix.
  const otherSuffix = String(accountSuffix) === '1213' ? '2224' : '1213';
  const selfWithOtherAccount = compact.includes('SELFFT') && compact.includes(otherSuffix);

  return explicitSelfTransfer || selfWithOtherAccount;
}

function classifyPaymentType_(paymentType) {
  const text = normalize_(paymentType).replace(/\s+/g, ' ');
  if (/g\s*[_-]?\s*pay/.test(text) || text.includes('google pay')) return 'GPAY';
  if (text.includes('cheque') || text.includes('check') || text.includes('chq')) return 'CHEQUE';
  if (text.includes('2224')) return 'DIRECT_2224';
  if (text.includes('1213')) return 'DIRECT_1213';
  return 'UNKNOWN';
}

function groupEntriesByDate_(entries, tz) {
  return entries.reduce((groups, entry) => {
    const key = dateKey_(entry.date, tz);
    if (!groups[key]) groups[key] = [];
    groups[key].push(entry);
    return groups;
  }, {});
}

function findCandidates_(transactions, amount, baseDate, fromDayOffset, toDayOffset, tz, options) {
  const allowedDates = buildAllowedDateSet_(baseDate, fromDayOffset, toDayOffset, tz);
  const includeBounced = Boolean(options && options.includeBounced);
  return transactions.filter(txn =>
    !txn.used &&
    !txn.isSelfAccount &&
    (includeBounced || !txn.isChequeBounced) &&
    allowedDates.has(txn.dateKey) &&
    Math.abs(txn.amount - amount) <= RECO_CONFIG.AMOUNT_TOLERANCE
  );
}

function findChequeCandidates_(transactions, entry, fromDayOffset, toDayOffset, tz) {
  const allowedDates = buildAllowedDateSet_(entry.date, fromDayOffset, toDayOffset, tz);
  return transactions
    .filter(txn =>
      !txn.used &&
      !txn.isSelfAccount &&
      allowedDates.has(txn.dateKey) &&
      Math.abs(txn.amount - entry.amount) <= RECO_CONFIG.AMOUNT_TOLERANCE &&
      transactionContainsChequeToken_(txn, entry.chequeTokens)
    )
    .sort((a, b) => a.date - b.date || a.accountName.localeCompare(b.accountName) || a.rowNumber - b.rowNumber);
}

function selectPreferredForwardDateCandidates_(candidates, baseDate) {
  if (!candidates.length) return [];
  const forwardOffsets = candidates.map(txn => daysBetween_(baseDate, txn.date));
  const minOffset = Math.min.apply(null, forwardOffsets);
  return candidates
    .filter(txn => daysBetween_(baseDate, txn.date) === minOffset)
    .sort((a, b) => a.rowNumber - b.rowNumber);
}

function selectNearestDateCandidates_(candidates, baseDate) {
  if (!candidates.length) return [];
  const distances = candidates.map(txn => Math.abs(daysBetween_(baseDate, txn.date)));
  const minDistance = Math.min.apply(null, distances);
  return candidates.filter(txn => Math.abs(daysBetween_(baseDate, txn.date)) === minDistance)
    .sort((a, b) => a.date - b.date || a.rowNumber - b.rowNumber);
}

function daysBetween_(dateA, dateB) {
  const a = new Date(dateA.getFullYear(), dateA.getMonth(), dateA.getDate(), 12, 0, 0);
  const b = new Date(dateB.getFullYear(), dateB.getMonth(), dateB.getDate(), 12, 0, 0);
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function buildAllowedDateSet_(baseDate, fromDayOffset, toDayOffset, tz) {
  const allowedDates = new Set();
  for (let i = fromDayOffset; i <= toDayOffset; i++) {
    allowedDates.add(dateKey_(addDays_(baseDate, i), tz));
  }
  return allowedDates;
}

function transactionContainsChequeToken_(txn, entryTokens) {
  // Numeric cheque numbers may lose leading zeroes when Google Sheets reads a
  // numeric CHQNO cell. For example, Vyapar can contain 000692 while the bank
  // CHQNO column contains 692. Compare canonical variants so both are equal.
  const txnCanonicalSet = new Set();
  (txn.chequeTokens || []).forEach(token =>
    chequeTokenVariants_(token).forEach(variant => txnCanonicalSet.add(variant))
  );

  const entryVariants = [];
  (entryTokens || []).forEach(token =>
    chequeTokenVariants_(token).forEach(variant => entryVariants.push(variant))
  );

  if (entryVariants.some(variant => txnCanonicalSet.has(variant))) return true;

  const rawNarration = String(txn.narration || '').toUpperCase();
  const narrationTokens = [];
  collectExplicitChequeTokens_(rawNarration, narrationTokens);
  collectClearingTokens_(rawNarration, narrationTokens);
  const narrationCanonicalSet = new Set();
  narrationTokens.forEach(token =>
    chequeTokenVariants_(token).forEach(variant => narrationCanonicalSet.add(variant))
  );

  return entryVariants.some(variant => narrationCanonicalSet.has(variant));
}

function chequeTokenVariants_(token) {
  const cleaned = String(token || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!cleaned) return [];

  const variants = new Set([cleaned]);
  if (/^\d+$/.test(cleaned)) {
    const withoutLeadingZeroes = cleaned.replace(/^0+(?=\d)/, '');
    variants.add(withoutLeadingZeroes || '0');

    // Six digits is the common cheque-number display format. This restores
    // zeroes stripped from numeric bank cells such as 692 -> 000692.
    if (withoutLeadingZeroes.length <= 6) {
      variants.add(withoutLeadingZeroes.padStart(6, '0'));
    }
  }
  return [...variants];
}

function extractVyaparChequeTokens_(refNo, paymentType) {
  // In VYAPAR_PAYMENT_IN the cheque number is normally stored in Payment Type,
  // for example: "Cheque (670873)". Prefer that number so an unrelated Ref No
  // cannot cause a wrong or duplicate cheque match.
  const paymentTypeTokens = [];
  collectExplicitChequeTokens_(String(paymentType || ''), paymentTypeTokens);
  collectChequeNumberInBrackets_(String(paymentType || ''), paymentTypeTokens);

  const rankedPaymentTypeTokens = rankChequeTokens_(paymentTypeTokens);
  if (rankedPaymentTypeTokens.length) return rankedPaymentTypeTokens;

  // Fallback only for older rows where the cheque number may be in Ref No.
  const refTokens = [];
  collectExplicitChequeTokens_(String(refNo || ''), refTokens);
  collectReferenceTokens_(String(refNo || ''), refTokens);
  return rankChequeTokens_(refTokens);
}

function extractBankChequeTokens_(chequeField, narration) {
  const tokens = [];
  collectBankChequeFieldTokens_(String(chequeField || ''), tokens);
  collectExplicitChequeTokens_(String(chequeField || ''), tokens);
  collectExplicitChequeTokens_(String(narration || ''), tokens);
  collectClearingTokens_(String(narration || ''), tokens);

  // Expand every token into leading-zero-safe variants before ranking.
  const expanded = [];
  tokens.forEach(token => chequeTokenVariants_(token).forEach(variant => expanded.push(variant)));
  return rankChequeTokens_(expanded);
}

function collectBankChequeFieldTokens_(text, output) {
  const upper = String(text || '').toUpperCase().trim();
  if (!upper || upper === '-') return;

  // Bank CHQNO columns are often numeric, so 000692 is returned by Sheets as
  // 692. Capture even short numeric values, then create a six-digit variant.
  const numericParts = upper.match(/\d{1,12}/g) || [];
  numericParts.forEach(token => {
    output.push(token);
    if (token.length <= 6) output.push(token.padStart(6, '0'));
  });

  const compact = upper.replace(/[^A-Z0-9]/g, '');
  if (/^[A-Z0-9]{1,12}$/.test(compact)) output.push(compact);
}

function collectExplicitChequeTokens_(text, output) {
  const upper = String(text || '').toUpperCase();
  // Supports formats such as:
  // Cheque 670873, Cheque No 670873, Cheque (670873), CHQ-670873.
  const regex = /(?:CHQ|CHEQUE|CHECK|INSTRUMENT)\s*(?:NO|NUMBER)?\s*[\s(:#./\-\[\{]*\s*([A-Z0-9]{4,12})/g;
  let match;
  while ((match = regex.exec(upper)) !== null) output.push(match[1]);
}

function collectChequeNumberInBrackets_(text, output) {
  const upper = String(text || '').toUpperCase();
  // Strictly captures the value in brackets after Cheque/CHQ.
  const regex = /(?:CHQ|CHEQUE|CHECK)\s*\(\s*([A-Z0-9]{4,12})\s*\)/g;
  let match;
  while ((match = regex.exec(upper)) !== null) output.push(match[1]);
}

function collectClearingTokens_(text, output) {
  const upper = String(text || '').toUpperCase();
  const regex = /(?:CLG|CLEARING)\s*[:#./-]*\s*([A-Z0-9]{4,12})/g;
  let match;
  while ((match = regex.exec(upper)) !== null) output.push(match[1]);
}

function collectReferenceTokens_(text, output) {
  const upper = String(text || '').toUpperCase().trim();
  if (!upper) return;

  const numericParts = upper.match(/\d{4,12}/g) || [];
  numericParts.forEach(token => output.push(token));

  const compact = upper.replace(/[^A-Z0-9]/g, '');
  if (/^[A-Z0-9]{4,12}$/.test(compact)) output.push(compact);
}

function rankChequeTokens_(tokens) {
  return [...new Set(tokens.map(token => String(token).toUpperCase()).filter(token => /^[A-Z0-9]{1,12}$/.test(token)))]
    .sort((a, b) => {
      const aScore = Math.abs(a.length - 6);
      const bScore = Math.abs(b.length - 6);
      return aScore - bScore || b.length - a.length || a.localeCompare(b);
    });
}

function buildPartyAwareSingleResult_(entry, candidates, defaultBankAccount, ruleText) {
  if (candidates.length === 1) {
    const txn = candidates[0];
    const similarity = partyNarrationSimilarity_(entry.party, txn.narration);
    return {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: txn.accountName,
      bankTxn: txn,
      difference: round2_(txn.amount - entry.amount),
      status: 'MATCHED',
      remarks: `Unique exact-amount bank credit found; party/narration similarity ${similarity}%.`
    };
  }

  if (candidates.length > 1) {
    const choice = chooseCandidateByPartySimilarity_(entry, candidates);
    if (choice.autoMatch) {
      return {
        entry,
        ruleText,
        groupTotal: '',
        bankAccount: choice.txn.accountName,
        bankTxn: choice.txn,
        difference: round2_(choice.txn.amount - entry.amount),
        status: 'MATCHED',
        remarks: `${candidates.length} exact-amount credits existed on the nearest date; party/narration similarity selected bank row ${choice.txn.rowNumber} (${choice.score}%, lead ${choice.lead} points).`
      };
    }

    const rankedText = choice.ranked
      .slice(0, 5)
      .map(item => `${item.txn.accountName}:${item.txn.rowNumber}=${item.score}%`)
      .join(', ');
    return {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: defaultBankAccount,
      bankTxn: null,
      difference: '',
      status: 'MANUAL CHECK',
      remarks: `${candidates.length} exact-amount credits were found on the nearest date, but narration did not identify one confidently. Similarity: ${rankedText}`
    };
  }

  return {
    entry,
    ruleText,
    groupTotal: '',
    bankAccount: defaultBankAccount,
    bankTxn: null,
    difference: '',
    status: 'UNMATCHED',
    remarks: 'No matching unused bank credit found.'
  };
}

function chooseCandidateByPartySimilarity_(entry, candidates) {
  const ranked = candidates
    .map(txn => ({
      txn,
      score: partyNarrationSimilarity_(entry.party, txn.narration)
    }))
    .sort((a, b) => b.score - a.score || a.txn.rowNumber - b.txn.rowNumber);

  if (!ranked.length) {
    return { autoMatch: false, txn: null, score: 0, lead: 0, ranked: [] };
  }

  const top = ranked[0];
  const secondScore = ranked.length > 1 ? ranked[1].score : 0;
  const lead = top.score - secondScore;
  const autoMatch =
    top.score >= RECO_CONFIG.PARTY_SIMILARITY_MIN_AUTO_MATCH &&
    lead >= RECO_CONFIG.PARTY_SIMILARITY_MIN_LEAD;

  return {
    autoMatch,
    txn: top.txn,
    score: top.score,
    lead,
    ranked
  };
}

function pairEntriesAndTransactionsByPartySimilarity_(entries, transactions) {
  const remainingEntries = entries.slice().sort((a, b) => a.rowNumber - b.rowNumber);
  const remainingTransactions = transactions.slice().sort((a, b) => a.rowNumber - b.rowNumber);
  const pairs = [];

  // First reserve positively identified party/narration pairs.
  while (remainingEntries.length && remainingTransactions.length) {
    let best = null;
    remainingEntries.forEach((entry, entryIndex) => {
      remainingTransactions.forEach((txn, txnIndex) => {
        const score = partyNarrationSimilarity_(entry.party, txn.narration);
        if (!best || score > best.score ||
            (score === best.score && txn.rowNumber < best.txn.rowNumber) ||
            (score === best.score && txn.rowNumber === best.txn.rowNumber && entry.rowNumber < best.entry.rowNumber)) {
          best = { entry, txn, entryIndex, txnIndex, score };
        }
      });
    });

    if (!best || best.score <= 0) break;
    pairs.push({ entry: best.entry, txn: best.txn, score: best.score });
    remainingEntries.splice(best.entryIndex, 1);
    remainingTransactions.splice(best.txnIndex, 1);
  }

  // When narration has no useful party text, preserve the original deterministic
  // row-order pairing instead of blocking otherwise valid exact-amount GPay matches.
  const fallbackCount = Math.min(remainingEntries.length, remainingTransactions.length);
  for (let index = 0; index < fallbackCount; index++) {
    pairs.push({
      entry: remainingEntries[index],
      txn: remainingTransactions[index],
      score: partyNarrationSimilarity_(remainingEntries[index].party, remainingTransactions[index].narration)
    });
  }

  return pairs;
}

function partyNarrationSimilarity_(party, narration) {
  const partyText = normalizeSimilarityText_(party);
  const narrationText = normalizeSimilarityText_(narration);
  if (!partyText || !narrationText) return 0;

  const partyCompact = partyText.replace(/\s+/g, '');
  const narrationCompact = narrationText.replace(/\s+/g, '');
  if (partyCompact.length >= 5 && narrationCompact.includes(partyCompact)) return 100;

  const partyTokens = similarityTokens_(partyText, true);
  const narrationTokens = similarityTokens_(narrationText, false);
  if (!partyTokens.length || !narrationTokens.length) return 0;

  const narrationSet = new Set(narrationTokens);
  let totalWeight = 0;
  let matchedWeight = 0;

  partyTokens.forEach(token => {
    const weight = similarityTokenWeight_(token);
    totalWeight += weight;

    if (narrationSet.has(token)) {
      matchedWeight += weight;
      return;
    }

    const closeToken = narrationTokens.find(other => tokensApproximatelyMatch_(token, other));
    if (closeToken) matchedWeight += weight * 0.75;
  });

  if (!totalWeight) return 0;
  const coverageScore = (matchedWeight / totalWeight) * 100;

  // Character bigram similarity helps with merged words, punctuation and minor
  // spelling differences commonly found in bank narrations.
  const charScore = diceCoefficient_(partyCompact, narrationCompact) * 100;
  return Math.max(0, Math.min(100, Math.round(Math.max(coverageScore, charScore))));
}

function normalizeSimilarityText_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function similarityTokens_(text, isParty) {
  const bankNoise = new Set([
    'NEFT', 'IMPS', 'RTGS', 'UPI', 'MOB', 'MOBILE', 'TRANSFER', 'TRF', 'FT',
    'BANK', 'AXIS', 'FEDERAL', 'HDFC', 'ICICI', 'SBI', 'PAYMENT', 'RECEIVED',
    'CREDIT', 'DEBIT', 'CR', 'DR', 'REF', 'REFERENCE', 'TXN', 'TRANSACTION',
    'CLG', 'CLEARING', 'CHEQUE', 'CHQ', 'BY', 'TO', 'FROM', 'OF', 'THE', 'IN',
    'AC', 'ACCOUNT', 'VIA', 'ONLINE', 'PURCHASE'
  ]);
  const legalNoise = new Set(['PRIVATE', 'PVT', 'LIMITED', 'LTD', 'LLP', 'OPC']);

  return text.split(' ').filter(token => {
    if (!token || token.length < 2 || /^\d+$/.test(token)) return false;
    if (bankNoise.has(token)) return false;
    if (isParty && legalNoise.has(token)) return false;
    return true;
  });
}

function similarityTokenWeight_(token) {
  const commonBusinessWords = new Set([
    'TRADERS', 'ENTERPRISES', 'CONSTRUCTION', 'CONSTRUCTIONS', 'SALES',
    'STEEL', 'CEMENT', 'HARDWARE', 'ELECTRICALS', 'ASSOCIATES', 'SERVICES',
    'DEVELOPERS', 'GRANITE', 'TRADING', 'INDUSTRIES', 'CORPORATION', 'COMPANY'
  ]);
  const base = Math.max(2, Math.min(10, token.length));
  return commonBusinessWords.has(token) ? base * 0.45 : base;
}

function tokensApproximatelyMatch_(a, b) {
  if (a === b) return true;
  if (a.length >= 4 && b.length >= 4 && (a.startsWith(b) || b.startsWith(a))) return true;
  if (Math.min(a.length, b.length) >= 5 && levenshteinDistance_(a, b) <= 1) return true;
  return false;
}

function levenshteinDistance_(a, b) {
  const previous = [];
  const current = [];
  for (let j = 0; j <= b.length; j++) previous[j] = j;

  for (let i = 1; i <= a.length; i++) {
    current[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      current[j] = Math.min(
        current[j - 1] + 1,
        previous[j] + 1,
        previous[j - 1] + cost
      );
    }
    for (let j = 0; j <= b.length; j++) previous[j] = current[j];
  }
  return previous[b.length];
}

function diceCoefficient_(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.length < 2 || b.length < 2) return 0;

  const pairs = new Map();
  for (let i = 0; i < a.length - 1; i++) {
    const pair = a.substring(i, i + 2);
    pairs.set(pair, (pairs.get(pair) || 0) + 1);
  }

  let intersection = 0;
  for (let i = 0; i < b.length - 1; i++) {
    const pair = b.substring(i, i + 2);
    const count = pairs.get(pair) || 0;
    if (count > 0) {
      pairs.set(pair, count - 1);
      intersection++;
    }
  }
  return (2 * intersection) / ((a.length - 1) + (b.length - 1));
}

function buildSingleResult_(entry, candidates, defaultBankAccount, ruleText) {
  if (candidates.length === 1) {
    const txn = candidates[0];
    return {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: txn.accountName,
      bankTxn: txn,
      difference: round2_(txn.amount - entry.amount),
      status: 'MATCHED',
      remarks: 'Unique bank credit found.'
    };
  }

  if (candidates.length > 1) {
    return {
      entry,
      ruleText,
      groupTotal: '',
      bankAccount: defaultBankAccount,
      bankTxn: null,
      difference: '',
      status: 'MANUAL CHECK',
      remarks: `${candidates.length} possible bank credits found. Rows: ${candidates.map(txn => `${txn.accountName}:${txn.rowNumber}`).join(', ')}`
    };
  }

  return {
    entry,
    ruleText,
    groupTotal: '',
    bankAccount: defaultBankAccount,
    bankTxn: null,
    difference: '',
    status: 'UNMATCHED',
    remarks: 'No matching unused bank credit found.'
  };
}

function markUsedIfUnique_(candidates, vyaparRowNumbers) {
  if (candidates.length === 1) {
    candidates[0].used = true;
    candidates[0].usedBy = vyaparRowNumbers.join(',');
  }
}

function writeRecoReport_(ss, entries, results) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.REPORT_SHEET);
  const headers = [
    'Vyapar Row', 'Vyapar Date', 'Ref No', 'Cheque No', 'Party', 'Entry Type',
    'Vyapar Amount', 'Payment Type', 'Matching Rule', 'GPay Group Total',
    'Matched Bank', 'Bank Row', 'Bank Date', 'Bank Narration',
    'Party/Narration Similarity %', 'Bank Cheque No', 'Bank Credit Amount',
    'Date Difference (Days)', 'Difference', 'Status', 'Remarks',
    'Bounce Debit Row', 'Bounce Date', 'Bounce Narration'
  ];

  const rows = entries
    .slice()
    .sort((a, b) => a.rowNumber - b.rowNumber)
    .map(entry => {
      const result = results.get(entry.rowNumber);
      const txn = result && result.bankTxn;
      const similarity = txn ? partyNarrationSimilarity_(entry.party, txn.narration) : '';
      const dayDifference = txn ? daysBetween_(entry.date, txn.date) : '';
      const amountBasis = result && result.groupTotal
        ? `Grouped GPay total ₹${Number(result.groupTotal).toFixed(2)}`
        : 'Exact amount';
      const evidence = txn
        ? `${amountBasis}; date difference ${dayDifference} day(s); party/narration similarity ${similarity}%`
        : '';
      const remarks = result
        ? [evidence, result.remarks].filter(Boolean).join(' | ')
        : 'No result created.';

      return [
        entry.rowNumber,
        entry.date,
        entry.refNo,
        formatChequeDisplay_(entry.chequeNo),
        entry.party,
        entry.entryType,
        entry.amount,
        entry.paymentType,
        result ? result.ruleText : '',
        result ? result.groupTotal : '',
        result ? result.bankAccount : '',
        txn ? txn.rowNumber : '',
        txn ? txn.date : '',
        txn ? txn.narration : '',
        similarity,
        txn ? formatChequeDisplay_(txn.chequeNo) : '',
        txn ? txn.amount : '',
        dayDifference,
        result ? result.difference : '',
        result ? result.status : 'MANUAL CHECK',
        remarks,
        txn && txn.isChequeBounced ? txn.bounceDebitRow : '',
        txn && txn.isChequeBounced ? txn.bounceDate : '',
        txn && txn.isChequeBounced ? txn.bounceNarration : ''
      ];
    });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    // Cheque numbers are identifiers, not money. Format as plain text before writing
    // so Google Sheets does not display a rupee sign or remove leading zeroes.
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 16, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  formatReportSheet_(sheet, rows.length, headers.length, [2, 13, 23], [7, 10, 17, 19]);
  if (rows.length) {
    sheet.getRange(2, 15, rows.length, 1).setNumberFormat('0"%"');
    sheet.getRange(2, 18, rows.length, 1).setNumberFormat('0');
  }
  applyStatusColors_(sheet, rows.length, 20);
}

function writeUnmatchedBankReport_(ss, transactions) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.UNMATCHED_BANK_SHEET);
  const headers = [
    'Bank Account', 'Source Sheet', 'Bank Row', 'Date', 'Narration',
    'Cheque No', 'Credit Amount', 'Status', 'Remarks'
  ];
  const rows = transactions
    .filter(txn => !txn.used && !txn.isSelfAccount && !txn.isChequeBounced)
    .sort((a, b) => a.accountName.localeCompare(b.accountName) || a.date - b.date || a.rowNumber - b.rowNumber)
    .map(txn => [
      txn.accountName,
      txn.sheetName,
      txn.rowNumber,
      txn.date,
      txn.narration,
      formatChequeDisplay_(txn.chequeNo),
      txn.amount,
      'UNMATCHED',
      'Unused bank credit; not classified as self-account or cheque-bounce.'
    ]);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  formatReportSheet_(sheet, rows.length, headers.length, [4], [7]);
  applyStatusColors_(sheet, rows.length, 8);
}


/**
 * Creates an import-ready list of bank transactions that are missing in Vyapar.
 *
 * Output columns intentionally follow the requested import layout:
 *   Vyapar Date | Type | Suggested Party Name | Amount | Bank Account Type | Bank Narration
 *
 * Payment Type is intentionally not included. Bank Account Type contains the exact
 * source bank sheet name (1213 or 2224), so the user can identify the account from
 * which each missing transaction came.
 *
 * Included rows:
 *   - Payment-In: unused bank credits
 *   - Payment-Out: eligible bank debits
 *
 * Excluded rows:
 *   - self-account transfers
 *   - cheque-bounce reversal debits
 *   - bounced credit presentations
 *   - bank credits already used in reconciliation
 *
 * Suggested Party Name is populated only when narration similarity is sufficiently
 * strong and clearly ahead of other party candidates. Otherwise it remains blank.
 */
function writeVyaparMissingEntriesExport_(ss, creditTransactions, debitTransactions, vyaparEntries) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.VYAPAR_MISSING_EXPORT_SHEET);
  const headers = [
    'Vyapar Date',
    'Type',
    'Suggested Party Name',
    'Amount',
    'Bank Account Type',
    'Bank Narration'
  ];
  const partyCandidates = buildUniquePartyCandidates_(vyaparEntries);

  const paymentInRows = creditTransactions
    .filter(txn => !txn.used && !txn.isSelfAccount && !txn.isChequeBounced)
    .map(txn => ({
      date: txn.date,
      narration: txn.narration,
      amount: txn.amount,
      bankAccountType: txn.sheetName || txn.accountName,
      accountName: txn.accountName,
      rowNumber: txn.rowNumber,
      entryType: 'Payment-In'
    }));

  const paymentOutRows = debitTransactions
    .filter(txn => !txn.isSelfAccount && !txn.isChequeBounceReversal)
    .map(txn => ({
      date: txn.date,
      narration: txn.narration,
      amount: txn.amount,
      bankAccountType: txn.sheetName || txn.accountName,
      accountName: txn.accountName,
      rowNumber: txn.rowNumber,
      entryType: 'Payment-Out'
    }));

  const rows = [...paymentInRows, ...paymentOutRows]
    .sort((a, b) =>
      a.date - b.date ||
      a.entryType.localeCompare(b.entryType) ||
      a.accountName.localeCompare(b.accountName) ||
      a.rowNumber - b.rowNumber
    )
    .map(txn => {
      const partyMatch = suggestPartyFromNarration_(txn.narration, partyCandidates);
      return [
        txn.date,
        txn.entryType,
        partyMatch ? partyMatch.party : '',
        txn.amount,
        txn.bankAccountType,
        txn.narration
      ];
    });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd-mm-yyyy');
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('#,##0.00');
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();

    // Light visual separation for review; values remain plain text for export.
    const entryTypes = sheet.getRange(2, 2, rows.length, 1).getValues();
    entryTypes.forEach((row, index) => {
      const cell = sheet.getRange(index + 2, 2);
      if (row[0] === 'Payment-In') cell.setBackground('#d9ead3');
      if (row[0] === 'Payment-Out') cell.setBackground('#fce5cd');
    });
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 115);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 260);
  sheet.setColumnWidth(6, 500);
}

function buildUniquePartyCandidates_(vyaparEntries) {
  const unique = new Map();
  vyaparEntries.forEach(entry => {
    const party = String(entry.party || '').trim().replace(/\s+/g, ' ');
    if (!party) return;
    const key = normalizeSimilarityText_(party);
    if (!key || unique.has(key)) return;
    unique.set(key, { party, key });
  });
  return [...unique.values()];
}

function suggestPartyFromNarration_(narration, partyCandidates) {
  if (!narration || !partyCandidates.length) return null;

  const ranked = partyCandidates
    .map(candidate => ({
      party: candidate.party,
      score: partyNarrationSimilarity_(candidate.party, narration)
    }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score || a.party.localeCompare(b.party));

  if (!ranked.length) return null;
  const best = ranked[0];
  const secondScore = ranked.length > 1 ? ranked[1].score : 0;
  const lead = best.score - secondScore;

  if (
    best.score >= RECO_CONFIG.MISSING_PARTY_SIMILARITY_MIN &&
    (ranked.length === 1 || lead >= RECO_CONFIG.MISSING_PARTY_SIMILARITY_MIN_LEAD)
  ) {
    return { party: best.party, score: best.score, lead };
  }
  return null;
}

function writeSelfAccountReport_(ss, transactions) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.SELF_ACCOUNT_SHEET);
  const headers = [
    'Bank Account', 'Source Sheet', 'Bank Row', 'Date', 'Narration',
    'Cheque/Reference No', 'Credit Amount', 'Status', 'Remarks'
  ];
  const rows = transactions
    .filter(txn => txn.isSelfAccount)
    .sort((a, b) => a.accountName.localeCompare(b.accountName) || a.date - b.date || a.rowNumber - b.rowNumber)
    .map(txn => [
      txn.accountName,
      txn.sheetName,
      txn.rowNumber,
      txn.date,
      txn.narration,
      formatChequeDisplay_(txn.chequeNo),
      txn.amount,
      'SELF ACCOUNT',
      'Own-account/internal transfer; excluded from reconciliation matching and unmatched-bank report.'
    ]);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  formatReportSheet_(sheet, rows.length, headers.length, [4], [7]);
  if (rows.length) {
    sheet.getRange(2, 8, rows.length, 1).setBackground('#cfe2f3').setFontWeight('bold');
  }
}

function writeBouncedChequeReport_(ss, transactions) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.BOUNCED_CHEQUE_SHEET);
  const headers = [
    'Bank Account', 'Source Sheet', 'Credit Row', 'Credit Date', 'Credit Narration',
    'Cheque No', 'Credit Amount', 'Reversal Debit Row', 'Reversal Date',
    'Reversal Narration', 'Reversal Amount', 'Matched Vyapar Row(s)', 'Status', 'Remarks'
  ];
  const rows = transactions
    .filter(txn => txn.isChequeBounced)
    .sort((a, b) => a.accountName.localeCompare(b.accountName) || a.date - b.date || a.rowNumber - b.rowNumber)
    .map(txn => [
      txn.accountName,
      txn.sheetName,
      txn.rowNumber,
      txn.date,
      txn.narration,
      formatChequeDisplay_(txn.chequeNo),
      txn.amount,
      txn.bounceDebitRow,
      txn.bounceDate,
      txn.bounceNarration,
      txn.bounceAmount,
      txn.usedBy,
      'CHEQUE BOUNCED',
      'Same-date cheque credit was reversed by a debit entry.'
    ]);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  formatReportSheet_(sheet, rows.length, headers.length, [4, 9], [7, 11]);
  applyStatusColors_(sheet, rows.length, 13);
}

function writePartyLedgerSummary_(ss, entries, results) {
  const sheet = getOrCreateCleanSheet_(ss, RECO_CONFIG.PARTY_SUMMARY_SHEET);
  const title = 'Party Ledger Summary - Vyapar and Bank Statement';
  const tableHeaderRow = 17;
  const headers = [
    'Party',
    'Vyapar Entries',
    'Bank-Linked Entries',
    'Vyapar Total Amount',
    'Matched Vyapar Amount',
    'Bank 1213 Matched Amount',
    'Bank 2224 Matched Amount',
    'Total Bank Statement Matched',
    'Unmatched Amount',
    'Manual Check Amount',
    'Cheque Bounced Amount',
    'Difference (Vyapar - Bank)',
    'First Entry Date',
    'Last Entry Date'
  ];

  const parties = new Map();

  entries.forEach(entry => {
    const cleanParty = String(entry.party || '').trim().replace(/\s+/g, ' ');
    const displayParty = cleanParty || '(Blank Party)';
    const partyKey = displayParty.toUpperCase();
    const result = results.get(entry.rowNumber);
    const status = result ? result.status : 'MANUAL CHECK';

    if (!parties.has(partyKey)) {
      parties.set(partyKey, {
        party: displayParty,
        vyaparCount: 0,
        bankLinkedCount: 0,
        total: 0,
        matched: 0,
        bank1213: 0,
        bank2224: 0,
        unmatched: 0,
        manual: 0,
        bounced: 0,
        firstDate: entry.date,
        lastDate: entry.date
      });
    }

    const item = parties.get(partyKey);
    item.vyaparCount += 1;
    item.total = round2_(item.total + entry.amount);

    if (status === 'MATCHED') {
      item.matched = round2_(item.matched + entry.amount);
      item.bankLinkedCount += 1;

      // For grouped GPay, allocate only the party's own Vyapar amount.
      const bankAccount = String(result && result.bankAccount || '');
      if (bankAccount.includes('1213')) {
        item.bank1213 = round2_(item.bank1213 + entry.amount);
      } else if (bankAccount.includes('2224')) {
        item.bank2224 = round2_(item.bank2224 + entry.amount);
      }
    } else if (status === 'UNMATCHED') {
      item.unmatched = round2_(item.unmatched + entry.amount);
    } else if (status === 'CHEQUE BOUNCED') {
      item.bounced = round2_(item.bounced + entry.amount);
    } else {
      item.manual = round2_(item.manual + entry.amount);
    }

    if (entry.date < item.firstDate) item.firstDate = entry.date;
    if (entry.date > item.lastDate) item.lastDate = entry.date;
  });

  const partyItems = [...parties.values()]
    .sort((a, b) => a.party.localeCompare(b.party, undefined, { sensitivity: 'base' }));

  const rows = partyItems.map(item => {
    const bankMatched = round2_(item.bank1213 + item.bank2224);
    const difference = round2_(item.total - bankMatched);
    return [
      item.party,
      item.vyaparCount,
      item.bankLinkedCount,
      item.total,
      item.matched,
      item.bank1213,
      item.bank2224,
      bankMatched,
      item.unmatched,
      item.manual,
      item.bounced,
      difference,
      item.firstDate,
      item.lastDate
    ];
  });

  const totalEntries = entries.length;
  const grandTotal = round2_(entries.reduce((sum, entry) => sum + entry.amount, 0));
  const bank1213Total = round2_(partyItems.reduce((sum, item) => sum + item.bank1213, 0));
  const bank2224Total = round2_(partyItems.reduce((sum, item) => sum + item.bank2224, 0));
  const bankMatchedTotal = round2_(bank1213Total + bank2224Total);
  const unmatchedTotal = round2_(partyItems.reduce((sum, item) => sum + item.unmatched, 0));
  const manualTotal = round2_(partyItems.reduce((sum, item) => sum + item.manual, 0));
  const bouncedTotal = round2_(partyItems.reduce((sum, item) => sum + item.bounced, 0));
  const differenceTotal = round2_(grandTotal - bankMatchedTotal);

  sheet.getRange(1, 1, 1, headers.length).merge();
  sheet.getRange(1, 1)
    .setValue(title)
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('left');

  const summaryRows = [
    ['Source sheet', RECO_CONFIG.SHEET_VYAPAR],
    ['Total parties', rows.length],
    ['Total Vyapar entries', totalEntries],
    ['Vyapar grand total', grandTotal],
    ['Matched in bank statements', bankMatchedTotal],
    ['Matched in Axis Bank-1213', bank1213Total],
    ['Matched in Axis Bank-2224', bank2224Total],
    ['Unmatched amount', unmatchedTotal],
    ['Manual check amount', manualTotal],
    ['Cheque bounced amount', bouncedTotal],
    ['Difference (Vyapar - Bank)', differenceTotal]
  ];

  sheet.getRange(3, 1, summaryRows.length, 2).setValues(summaryRows);
  sheet.getRange(3, 1, summaryRows.length, 1)
    .setFontWeight('bold')
    .setBackground('#e7e6e6');
  sheet.getRange(6, 2, 8, 1).setNumberFormat('₹#,##0.00');

  sheet.getRange(15, 1, 1, headers.length).merge();
  sheet.getRange(15, 1)
    .setValue("Note: Bank amounts are party-wise amounts linked through reconciliation. Grouped GPay credits are allocated using each party's Vyapar amount. Unmatched bank credits cannot be assigned to a party automatically.")
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setWrap(true);

  sheet.getRange(tableHeaderRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(tableHeaderRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setWrap(true);

  if (rows.length) {
    sheet.getRange(tableHeaderRow + 1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(tableHeaderRow + 1, 4, rows.length, 9).setNumberFormat('₹#,##0.00');
    sheet.getRange(tableHeaderRow + 1, 13, rows.length, 2).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(tableHeaderRow, 1, rows.length + 1, headers.length).createFilter();

    const totalRow = tableHeaderRow + rows.length + 1;
    sheet.getRange(totalRow, 1).setValue('GRAND TOTAL');
    for (let col = 2; col <= 12; col++) {
      const letter = columnLetter_(col);
      sheet.getRange(totalRow, col).setFormula(`=SUM(${letter}${tableHeaderRow + 1}:${letter}${totalRow - 1})`);
    }
    sheet.getRange(totalRow, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#d9eaf7');
    sheet.getRange(totalRow, 4, 1, 9).setNumberFormat('₹#,##0.00');
  }

  sheet.setFrozenRows(tableHeaderRow);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidths(2, 2, 115);
  sheet.setColumnWidths(4, 9, 155);
  sheet.setColumnWidths(13, 2, 120);
}

function columnLetter_(columnNumber) {
  let number = columnNumber;
  let letter = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    number = Math.floor((number - 1) / 26);
  }
  return letter;
}

function formatReportSheet_(sheet, dataRowCount, columnCount, dateColumns, amountColumns) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');

  if (dataRowCount > 0) {
    sheet.getRange(1, 1, dataRowCount + 1, columnCount).createFilter();
    dateColumns.forEach(col => sheet.getRange(2, col, dataRowCount, 1).setNumberFormat('dd-mmm-yyyy'));
    amountColumns.forEach(col => sheet.getRange(2, col, dataRowCount, 1).setNumberFormat('₹#,##0.00'));
  }

  sheet.autoResizeColumns(1, columnCount);
  if (columnCount >= 14) sheet.setColumnWidth(14, 320);
  sheet.setColumnWidth(columnCount, 420);
}

function applyStatusColors_(sheet, dataRowCount, statusColumn) {
  if (!dataRowCount) return;
  const range = sheet.getRange(2, statusColumn, dataRowCount, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MATCHED').setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('UNMATCHED').setBackground('#f4cccc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('CHEQUE BOUNCED').setBackground('#e06666').setFontColor('#ffffff').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MANUAL CHECK').setBackground('#fff2cc').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function getRequiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet not found: ${name}`);
  return sheet;
}

function getOrCreateCleanSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();
  sheet.setConditionalFormatRules([]);
  return sheet;
}

function findBankHeaderRow_(data, aliases) {
  const scanCount = Math.min(RECO_CONFIG.HEADER_SCAN_ROWS, data.length);
  let bestIndex = -1;
  let bestScore = -1;

  for (let r = 0; r < scanCount; r++) {
    const headers = data[r];
    const hasDate = findColumn_(headers, aliases.date) !== -1;
    const hasCreditOrAmount =
      findColumn_(headers, aliases.credit) !== -1 ||
      findColumn_(headers, aliases.amount) !== -1;

    if (!hasDate || !hasCreditOrAmount) continue;

    let score = 0;
    Object.keys(aliases).forEach(field => {
      if (findColumn_(headers, aliases[field]) !== -1) score++;
    });

    if (score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }

  if (bestIndex === -1) {
    throw new Error(`Bank header row not detected in the first ${scanCount} rows. Date and Credit/Amount headings are required.`);
  }
  return bestIndex;
}

function findHeaderRow_(data, aliases, requiredFields) {
  const scanCount = Math.min(RECO_CONFIG.HEADER_SCAN_ROWS, data.length);
  let bestIndex = -1;
  let bestScore = -1;

  for (let r = 0; r < scanCount; r++) {
    const headers = data[r];
    let score = 0;
    Object.keys(aliases).forEach(field => {
      if (findColumn_(headers, aliases[field]) !== -1) score++;
    });

    const requiredFound = requiredFields.every(field => findColumn_(headers, aliases[field]) !== -1);
    if (requiredFound && score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }

  if (bestIndex === -1) {
    throw new Error(`Header row not detected. Required fields: ${requiredFields.join(', ')}`);
  }
  return bestIndex;
}

function mapColumns_(headers, aliases, fields, sheetName) {
  const result = {};
  fields.forEach(field => {
    result[field] = findColumn_(headers, aliases[field]);
    if (result[field] === -1) {
      throw new Error(`Column '${field}' not found in ${sheetName}. Headers found: ${headers.join(' | ')}`);
    }
  });
  return result;
}

function findColumn_(headers, aliases) {
  const normalHeaders = headers.map(normalize_);
  const normalAliases = aliases.map(normalize_);

  for (const alias of normalAliases) {
    const exactIndex = normalHeaders.indexOf(alias);
    if (exactIndex !== -1) return exactIndex;
  }

  for (let i = 0; i < normalHeaders.length; i++) {
    for (const alias of normalAliases) {
      if (normalHeaders[i] && (normalHeaders[i].includes(alias) || alias.includes(normalHeaders[i]))) return i;
    }
  }
  return -1;
}

function normalize_(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[_\-]+/g, ' ')
    .replace(/[^a-z0-9/ ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseAmount_(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '') return NaN;

  const text = String(value).trim();
  const isNegative = /^\s*-/.test(text) || /\((.*)\)/.test(text) || /\bdr\b/i.test(text);
  const cleaned = text.replace(/[^0-9.]/g, '');
  if (!cleaned) return NaN;
  const number = Number(cleaned);
  return isNegative ? -number : number;
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 12, 0, 0);
  }
  if (value === null || value === undefined || value === '') return null;

  const text = String(value).trim();
  let match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(match[2]) - 1, Number(match[1]), 12, 0, 0);
  }

  match = text.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 12, 0, 0);
  }
  return null;
}

function addDays_(date, days) {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  result.setDate(result.getDate() + days);
  return result;
}

function dateKey_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function amountKey_(amount) {
  return round2_(amount).toFixed(2);
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatChequeDisplay_(value) {
  if (value === null || value === undefined || value === '') return '';
  return String(value).trim();
}

function round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}
