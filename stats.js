// stats.js - EXACTE EXCEL VBA MACRO KLASSEMENT OPMAAK
// ==================== INITIALISATIE ====================
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ stats.js geladen - Exacte Excel VBA macro opmaak');
    
    const currentPage = document.querySelector('.page.active');
    if (currentPage && currentPage.id === 'page8') {
        loadStatsPage();
    }
    if (currentPage && currentPage.id === 'page9') {
        loadRankingPage();
    }
});

// ==================== DATA FUNCTIES ====================
function getPlayers() {
    let players = [];
    const savedState = localStorage.getItem('billiardState');
    
    if (savedState) {
        try {
            const state = JSON.parse(savedState);
            players = state.players || [];
        } catch (e) {
            console.error('❌ Fout bij laden billiardState:', e);
        }
    }
    
    if (players.length === 0) {
        players = JSON.parse(localStorage.getItem('biljartPlayers') || '[]');
    }
    
    return players;
}

function getMatches() {
    let allMatches = [];
    const savedState = localStorage.getItem('billiardState');
    
    if (savedState) {
        try {
            const state = JSON.parse(savedState);
            allMatches = state.matches || [];
        } catch (e) {
            console.error('❌ Fout bij laden matches:', e);
        }
    }
    
    // Filter alleen VOLTOOIDE matches
    const completedMatches = allMatches.filter(m => m.completed === true);
    console.log(`🏁 ${completedMatches.length} voltooide matches voor klassement`);
    
    return completedMatches;
}

function getSeizoen() {
    const savedState = localStorage.getItem('billiardState');
    if (savedState) {
        try {
            const state = JSON.parse(savedState);
            if (state.seizoen) return state.seizoen;
        } catch (e) {}
    }
    
    const players = getPlayers();
    if (players.length > 0 && players[0].seizoen) {
        return players[0].seizoen;
    }
    
    return '2024-2025';
}

// ==================== EXACTE MACRO LOGICA ====================
function berekenExactMacroKlassement() {
    const players = getPlayers();
    const matches = getMatches();
    
    // Haal speeldagen uit state (geïmporteerd via CSV)
    let speeldagenData = [];
    const savedState = localStorage.getItem('billiardState');
    if (savedState) {
        try {
            const state = JSON.parse(savedState);
            if (state.speeldagen && Array.isArray(state.speeldagen)) {
                speeldagenData = state.speeldagen;
            }
        } catch (e) {
            console.error('Fout bij laden speeldagen uit state:', e);
        }
    }
    
    // Als er GEEN speeldagen geïmporteerd zijn, fallback naar bestaande logica
    if (speeldagenData.length === 0) {
        if (matches.length === 0) {
            return { speeldagen: [], spelersKlassement: [] };
        }
        const alleDatums = matches.map(m => m.date);
        const uniekeDatums = [...new Set(alleDatums)].sort((a, b) => new Date(a) - new Date(b));
        speeldagenData = uniekeDatums.map((datum, index) => ({
            datum: datum,
            speeldagNummer: index + 1,
            jsDatum: new Date(datum),
            displayDatum: datum
        }));
    }
    
    // Formatteer naar het formaat dat de rest van de code verwacht
    const speeldagen = speeldagenData.map((item, index) => {
        const d = new Date(item.jsDatum || item.datum);
        const dag = d.getDate();
        const maanden = ['jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 
                        'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
        const maand = maanden[d.getMonth()];
        return {
            datum: item.displayDatum || item.datum,
            speeldagNummer: item.volgnummer || (index + 1),
            datumShort: `${dag}-${maand}`
        };
    });
    
    // Sorteer op datum (voor het geval de CSV niet gesorteerd is)
    speeldagen.sort((a, b) => new Date(a.datum) - new Date(b.datum));
    
    // Bereken matchpunten per speler per speeldag
    const spelersMetPunten = players.filter(p => p.name).map(player => {
        const spelerNaam = player.name;
        const spelerMatches = matches.filter(m => 
            (m.p1 === spelerNaam || m.p2 === spelerNaam) ||
            (m.originalP1 === spelerNaam || m.originalP2 === spelerNaam)
        );
        
        const puntenPerDatum = {};
        speeldagen.forEach(speeldag => {
            puntenPerDatum[speeldag.datum] = [];
        });
        
        spelerMatches.forEach(match => {
            if (puntenPerDatum[match.date]) {
                const isWinnaar = match.winner === spelerNaam;
                const matchpunten = isWinnaar ? 2 : 1;
                puntenPerDatum[match.date].push(matchpunten);
            }
        });
        
        // Bereken tussentotalen
        const tussentotalen = [];
        let lopendTotaal = 0;
        
        speeldagen.forEach(speeldag => {
            const puntenDezeDag = puntenPerDatum[speeldag.datum] || [];
            const somDezeDag = puntenDezeDag.reduce((a, b) => a + b, 0);
            lopendTotaal += somDezeDag;
            tussentotalen.push(lopendTotaal);
        });
        
        return {
            naam: spelerNaam,
            puntenPerDatum: puntenPerDatum,
            tussentotalen: tussentotalen,
            eindtotaal: lopendTotaal
        };
    });
    
    // Sorteer op eindtotaal
    const gesorteerdeSpelers = [...spelersMetPunten].sort((a, b) => {
        if (b.eindtotaal !== a.eindtotaal) {
            return b.eindtotaal - a.eindtotaal;
        }
        return a.naam.localeCompare(b.naam);
    });
    
    gesorteerdeSpelers.forEach((speler, index) => {
        speler.positie = index + 1;
    });
    
    return {
        speeldagen: speeldagen,
        spelersKlassement: gesorteerdeSpelers
    };
}

// ==================== EXACTE MACRO KLASSEMENT PAGINA ====================
function loadRankingPage() {
    console.log('🏆 loadRankingPage() - Exacte Excel VBA Macro');
    
    try {
        const klassement = berekenExactMacroKlassement();
        const { speeldagen, spelersKlassement } = klassement;
        const seizoen = getSeizoen();
        
        if (spelersKlassement.length === 0) {
            return showNoDataMessage('rankingContent', 'voltooide matches');
        }
        
        const aantalSpeeldagen = speeldagen.length;
        const totaleKolommen = 2 + (aantalSpeeldagen * 3) + 1;
        
        // START HTML BOUWEN - EXACT ZOALS EXCEL MACRO
        let html = `
            <div class="excel-macro-klassement" style="font-family: 'Calibri', Arial, sans-serif;">
                <!-- TITEL RIJ (RIJ 1) - LICHTBLAUW ACHTERGROND, RODE TEKST -->
                <table style="width: ${totaleKolommen * 40 + 100}px; border-collapse: collapse; margin-bottom: 20px;">
                    <tr>
                        <td colspan="${totaleKolommen}" 
                            style="height: 37.5px; background: #4472c4; color: #c00000; font-family: 'Comic Sans MS'; 
                                   font-size: 24px; font-weight: bold; text-align: center; vertical-align: middle; border: 1px solid #000000;">
                            Klassement Biljartclub SPORT & VERMAAK ${seizoen}
                        </td>
                    </tr>
                </table>
                
                <!-- HOOFDTABEL -->
                <div style="overflow-x: auto; border: 1px solid #000000;">
                    <table style="border-collapse: collapse; min-width: ${totaleKolommen * 40}px;">
                        <!-- KOPTEKST 2 RIJEN -->
                        <thead>
                            <!-- RIJ 2: Naam + Speeldagnummers -->
                            <tr>
                                <!-- NAAM KOLOM (B) - ACHTERGROND #E2EFDA -->
                                <td rowspan="2" 
                                    style="border: 1px solid #000000; background: #E2EFDA; color: #000000; font-weight: bold; 
                                           text-align: center; vertical-align: bottom; white-space: nowrap; padding: 2px 8px; min-width: 120px;">
                                    Naam
                                </td>
        `;
        
        // PER SPEELDAG: SPEELDAGNUMMER (2 kolommen) + TOTAAL KOLOM
        speeldagen.forEach((speeldag, index) => {
            // Speeldagnummer over 2 kolommen
            html += `
                <!-- Speeldag ${speeldag.speeldagNummer} - Nummer -->
                <td colspan="2" 
                    style="border: 1px solid #000000; background: #E2EFDA; color: #000000; font-weight: bold; 
                           text-align: center; vertical-align: middle; padding: 2px; min-width: 20px;">
                    ${speeldag.speeldagNummer}
                </td>
                
                <!-- Tussentotaal kolom voor deze speeldag - TEKST LICHTBLAUW -->
                <td rowspan="2"
                    style="border: 1px solid #000000; background: #E2EFDA; color: #0078d4; font-weight: bold; 
                           writing-mode: vertical-lr; text-align: center; vertical-align: middle;
                           transform: rotate(180deg); padding: 2px; min-width: 15px;">
                    totaal
                </td>
            `;
        });
        
        // LAATSTE TOTAAL KOLOM (EINDTOTAAL) - TEKST LICHTBLAUW
        html += `
                                <!-- Eindtotaal kolom - TEKST LICHTBLAUW -->
                                <td rowspan="2"
                                    style="border: 1px solid #000000; background: #E2EFDA; color: #0078d4; font-weight: bold; 
                                           writing-mode: vertical-lr; text-align: center; vertical-align: middle;
                                           transform: rotate(180deg); padding: 2px; min-width: 20px;">
                                    totaal
                                </td>
                            </tr>
                            
                            <!-- RIJ 3: Datums -->
                            <tr>
        `;
        
        // DATUM KOLOMMEN (2 per speeldag)
        speeldagen.forEach((speeldag, index) => {
            html += `
                <!-- Datum voor speeldag ${speeldag.speeldagNummer} -->
                <td colspan="2"
                    style="border: 1px solid #000000; background: #E2EFDA; color: #000000; 
                           font-size: 11px; text-align: center; vertical-align: middle; padding: 2px;">
                    ${speeldag.datumShort}
                </td>
            `;
        });
        
        html += `
                            </tr>
                        </thead>
                        
                        <!-- SPELER RIJEN -->
                        <tbody>
        `;
        
        // SPELER DATA
        spelersKlassement.forEach((speler, spelerIndex) => {
            html += `
                <tr>
                    <!-- SPELER NAAM - ACHTERGROND #DDEBF7 -->
                    <td style="border: 1px solid #000000; background: #DDEBF7; padding: 2px 8px; font-weight: bold; white-space: nowrap; color: #000000;">
                        ${speler.positie}. ${speler.naam}
                    </td>
            `;
            
            // MATCHPUNTEN PER SPEELDAG (2 kolommen per speeldag)
            speeldagen.forEach((speeldag, dagIndex) => {
                const punten = speler.puntenPerDatum[speeldag.datum] || [];
                
                // Bepaal achtergrond voor matchkolommen (alternerend per datum)
                // Oneven speeldag: #FFF2CC, Even speeldag: wit
                const isOnevenDatum = (dagIndex % 2 === 0); // 0-based index
                const matchAchtergrond = isOnevenDatum ? '#FFF2CC' : '#ffffff';
                
                // Match 1
                const match1Punten = punten[0] || '';
                html += `
                    <td style="border: 1px solid #000000; background: ${matchAchtergrond}; text-align: center; padding: 2px; min-width: 20px; color: #000000;">
                        ${match1Punten}
                    </td>
                `;
                
                // Match 2
                const match2Punten = punten[1] || '';
                html += `
                    <td style="border: 1px solid #000000; background: ${matchAchtergrond}; text-align: center; padding: 2px; min-width: 20px; color: #000000;">
                        ${match2Punten}
                    </td>
                `;
                
                // TUSSENTOTAAL voor deze speeldag - ACHTERGROND #F8CBAD, TEKST LICHTBLAUW
                const tussentotaal = speler.tussentotalen[dagIndex] || 0;
                html += `
                    <td style="border: 1px solid #000000; background: #F8CBAD; font-weight: bold; color: #0078d4; text-align: center; padding: 2px; min-width: 15px;">
                        ${tussentotaal}
                    </td>
                `;
            });
            
            // EINDTOTAAL KOLOM - ACHTERGROND #F4DBBC, RODE TEKST
            html += `
                    <td style="border: 1px solid #000000; background: #F4DBBC; font-weight: bold; color: #c00000; text-align: center; padding: 4px; min-width: 20px;">
                        ${speler.eindtotaal}
                    </td>
                </tr>
            `;
        });
        
        // SLUIT TABEL
        html += `
                        </tbody>
                    </table>
                </div>
                
                <!-- KLEUREN LEGENDA VERWIJDERD UIT RANGSCHIKKING -->
                
                <!-- PRINT OPTIMALISATIE -->
                <style>
                    @media print {
                        .excel-macro-klassement {
                            zoom: 85%;
                        }
                        
                        table {
                            page-break-inside: auto !important;
                        }
                        
                        tr {
                            page-break-inside: avoid !important;
                            page-break-after: auto !important;
                        }
                        
                        /* Verberg knoppen bij printen */
                        button {
                            display: none !important;
                        }
                    }
                </style>
            </div>
        `;
        
        const content = document.getElementById('rankingContent');
        if (content) {
            content.innerHTML = html;
            console.log('✅ Exacte Excel VBA Macro Klassement getoond');
        }
        
    } catch (error) {
        console.error('❌ Fout in loadRankingPage:', error);
        showErrorMessage('rankingContent', error);
    }
}

// ==================== SPELER STATISTIEKEN PAGINA (ORIGINEEL) ====================
function loadStatsPage() {
    console.log('📈 loadStatsPage()');
    
    try {
        const players = getPlayers();
        const matches = getMatches();
        
        if (players.length === 0) {
            return showNoDataMessage('statsContent', 'spelers');
        }
        
        let html = `
            <div class="stats-header">
                <h2>📊 Speler Statistieken</h2>
                <p>${players.length} spelers, ${matches.length} voltooide matches</p>
            </div>
            <div class="players-stats-grid">
        `;
        
        players.forEach(player => {
            const spelerNaam = player.name;
            if (!spelerNaam) return;
            
            const spelerMatches = matches.filter(m => 
                m.p1 === spelerNaam || m.p2 === spelerNaam
            );
            
            if (spelerMatches.length === 0) {
                html += `
                    <div class="player-stats-card no-data-card">
                        <h3>${spelerNaam}</h3>
                        <p class="no-data-text">Geen voltooide matches</p>
                        <div class="stats-summary">
                            <p><strong>TSG:</strong> ${player.tsg || 'N/A'}</p>
                            <p><strong>Target:</strong> ${player.target || 'N/A'}</p>
                        </div>
                        <button class="detail-btn disabled-btn" disabled>
                            📈 Geen data
                        </button>
                    </div>
                `;
                return;
            }
            
            const gewonnenMatches = spelerMatches.filter(m => m.winner === spelerNaam);
            const winPercentage = (gewonnenMatches.length / spelerMatches.length * 100).toFixed(1);
            
            // Bereken gemiddelde score
            let totaalScore = 0;
            let totaalBeurten = 0;
            
            spelerMatches.forEach(match => {
                const isPlayer1 = match.p1 === spelerNaam;
                totaalScore += isPlayer1 ? (match.p1Score || 0) : (match.p2Score || 0);
                if (isPlayer1 && match.p1Turns) {
                    totaalBeurten += match.p1Turns.length;
                } else if (!isPlayer1 && match.p2Turns) {
                    totaalBeurten += match.p2Turns.length;
                }
            });
            
            const gemiddelde = totaalBeurten > 0 ? (totaalScore / totaalBeurten).toFixed(2) : '0.00';
            
            // Hoogste reeks
            let hoogsteReeks = 0;
            spelerMatches.forEach(match => {
                const isPlayer1 = match.p1 === spelerNaam;
                const reeks = isPlayer1 ? (match.p1Highest || 0) : (match.p2Highest || 0);
                if (reeks > hoogsteReeks) hoogsteReeks = reeks;
            });
            
            html += `
                <div class="player-stats-card">
                    <div class="card-header-excel">
                        <h3>${spelerNaam}</h3>
                        <span class="match-count-badge">${spelerMatches.length} matchen</span>
                    </div>
                    
                    <button onclick="showSpelerDetail('${spelerNaam}')" 
                            class="detail-btn">
                        📈 Detail Statistieken
                    </button>
                    
                    <div class="stats-summary">
                        <p><strong>Gespeeld:</strong> ${spelerMatches.length}</p>
                        <p><strong>Gewonnen:</strong> ${gewonnenMatches.length}</p>
                        <p><strong>Win %:</strong> ${winPercentage}%</p>
                        <p><strong>Gemiddelde:</strong> ${gemiddelde}</p>
                        <p><strong>Hoogste reeks:</strong> ${hoogsteReeks}</p>
                        <p><strong>TSG:</strong> ${player.tsg || 'N/A'}</p>
                    </div>
                </div>
            `;
        });
        
        html += `</div>`;
        
        const content = document.getElementById('statsContent');
        if (content) {
            content.innerHTML = html;
            console.log('✅ Statistieken getoond');
        }
        
    } catch (error) {
        console.error('❌ Fout in loadStatsPage:', error);
        showErrorMessage('statsContent', error);
    }
}

// ==================== SPELER DETAIL FUNCTIES (ORIGINEEL) ====================
function generateExcelStylePlayerDetail(playerName) {
    const players = getPlayers();
    const matches = getMatches();
    
    const player = players.find(p => p.name === playerName);
    if (!player) {
        return '<div class="error-message">Speler niet gevonden</div>';
    }
    
    const spelerMatches = matches.filter(m => 
        m.p1 === playerName || m.p2 === playerName
    );
    
    if (spelerMatches.length === 0) {
        return `
            <div class="no-data-message">
                <h3>Geen match geschiedenis</h3>
                <p>${playerName} heeft nog geen voltooide matches.</p>
            </div>
        `;
    }
    
    // Sorteer matches op datum
    spelerMatches.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    const targetEersteZes = player.targetEersteZes || 0;
    const tsgEersteZes = player.tsgEersteZes || 0;
    const tsg = player.tsg ? parseFloat(player.tsg.replace(',', '.')) : 0;
    const target = player.target || 50;
    
    let totaalGewonnen = 0;
    let cumulatieveJ = 0;
    let resultRows = [];
    
    spelerMatches.forEach((match, index) => {
        const isPlayer1 = match.p1 === playerName;
        const tegenstander = isPlayer1 ? match.p2 : match.p1;
        const punten = isPlayer1 ? (match.p1Score || 0) : (match.p2Score || 0);
        const beurten = isPlayer1 ? 
            (match.p1Turns ? match.p1Turns.length : 0) : 
            (match.p2Turns ? match.p2Turns.length : 0);
        const hoogsteReeks = isPlayer1 ? (match.p1Highest || 0) : (match.p2Highest || 0);
        const gewonnen = match.winner === playerName;
        
        const aantalPartijen = index + 1;
        if (gewonnen) totaalGewonnen++;
        
        const gemiddeldePerPartij = beurten > 0 ? (punten / beurten) : 0;
        cumulatieveJ += gemiddeldePerPartij;
        const totaalGemiddelde = cumulatieveJ / aantalPartijen;
        
        let winstVerlies;
        if (aantalPartijen <= 6 && tsgEersteZes > 0) {
            winstVerlies = ((totaalGemiddelde / tsgEersteZes) * 100 - 100);
        } else {
            winstVerlies = ((totaalGemiddelde / tsg) * 100 - 100);
        }
        
        let percHoogsteReeks;
        if (aantalPartijen <= 6 && targetEersteZes > 0) {
            percHoogsteReeks = (hoogsteReeks / targetEersteZes) * 100;
        } else {
            percHoogsteReeks = (hoogsteReeks / target) * 100;
        }
        
        const percGewonnen = (totaalGewonnen / aantalPartijen) * 100;
        const percVerloren = 100 - percGewonnen;
        const matchpunten = gewonnen ? 2 : 1;
        
        resultRows.push({
            speeldag: Math.floor(index / 2) + 1,
            tegenstander: tegenstander,
            datum: match.date,
            punten: punten,
            beurten: beurten,
            hoogsteReeks: hoogsteReeks,
            aantalPartijen: aantalPartijen,
            gewonnen: gewonnen ? "Ja" : "Nee",
            totaalGewonnen: totaalGewonnen,
            gemiddeldePerPartij: gemiddeldePerPartij,
            totaalGemiddelde: totaalGemiddelde,
            winstVerlies: winstVerlies,
            percHoogsteReeks: percHoogsteReeks,
            percGewonnen: percGewonnen,
            percVerloren: percVerloren,
            matchpunten: matchpunten
        });
    });
    
    return createExcelTableHTML(playerName, player, resultRows);
}

function formatDateForKlassement(dateStr) {
    if (!dateStr) return '';
    try {
        const d = new Date(dateStr);
        const dag = d.getDate();
        const maanden = ['jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 
                        'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
        const maand = maanden[d.getMonth()];
        return `${dag} ${maand}`;
    } catch (error) {
        return dateStr;
    }
}

function createExcelTableHTML(playerName, player, rows) {
    const lastRow = rows[rows.length - 1];
    const targetEersteZes = player.targetEersteZes || '';
    const tsgEersteZes = player.tsgEersteZes || '';
    
    const totaalPunten = rows.reduce((sum, row) => sum + row.punten, 0);
    const totaalBeurten = rows.reduce((sum, row) => sum + row.beurten, 0);
    const maxHoogsteReeks = rows.reduce((max, row) => Math.max(max, row.hoogsteReeks), 0);
    const totaalMatchpunten = rows.reduce((sum, row) => sum + row.matchpunten, 0);
    
    let html = `
        <div class="excel-table-container">
            <div class="table-header">
                <h3>${playerName} - Excel Stijl Overzicht</h3>
                <button class="small-btn" onclick="exportPlayerExcel('${playerName}')">📥 Exporteer</button>
            </div>
            
            <div class="table-wrapper">
                <table class="excel-table">
                    <thead class="excel-header">
                        <tr>
                            <th colspan="3" style="text-align: center; background: #c8e6c9; color: #000;">${playerName}</th>
                            <th colspan="3" style="text-align: center; background: #c8e6c9; color: #000;">Gemiddelde eerste 6 matchen</th>
                            <th style="background: #f4dbbc; text-align: center;">${targetEersteZes}</th>
                            <th colspan="3" style="text-align: center; background: #c8e6c9; color: #000;">Te maken punten eerste 6 matchen</th>
                            <th style="background: #f4dbbc; text-align: center;">${tsgEersteZes}</th>
                            <th style="background: #c8e6c9; text-align: center;">Gemiddelde</th>
                            <th style="background: #ffcccc; text-align: center; color: #d00;">1,00</th>
                            <th style="background: #c8e6c9; text-align: center;">Te maken punten</th>
                            <th style="background: #ffcccc; text-align: center; color: #d00;">20</th>
                            <th rowspan="2" style="background: #c8e6c9; vertical-align: middle; text-align: center;">Behaalde<br>matchpunten</th>
                        </tr>
                        
                        <tr class="excel-subheader">
                            <th>Speeldag</th>
                            <th>Tegenspeler</th>
                            <th>Datum</th>
                            <th>Punten</th>
                            <th>Beurten</th>
                            <th>Hoogste Reeks</th>
                            <th>Aantal<br>Partijen</th>
                            <th>Gewonnen</th>
                            <th>Totaal<br>Gewonnen</th>
                            <th>Gemiddelde<br>per Partij</th>
                            <th>Totaal<br>Gemiddelde</th>
                            <th>% Winst<br>Verlies</th>
                            <th>% Hoogste<br>Reeks</th>
                            <th>%<br>Gewonnen</th>
                            <th>%<br>Verloren</th>
                        </tr>
                    </thead>
                    
                    <tbody>
    `;
    
    rows.forEach((row, index) => {
        const rowClass = index % 2 === 0 ? 'excel-even' : 'excel-odd';
        const gewonnenClass = row.gewonnen === "Ja" ? 'excel-won' : 'excel-lost';
        
        html += `
            <tr class="${rowClass} ${gewonnenClass}">
                <td class="excel-cell number">${row.speeldag}</td>
                <td class="excel-cell">${row.tegenstander}</td>
                <td class="excel-cell">${formatDateForKlassement(row.datum)}</td>
                <td class="excel-cell number">${row.punten}</td>
                <td class="excel-cell number">${row.beurten}</td>
                <td class="excel-cell number">${row.hoogsteReeks}</td>
                <td class="excel-cell number">${row.aantalPartijen}</td>
                <td class="excel-cell">${row.gewonnen}</td>
                <td class="excel-cell number">${row.totaalGewonnen}</td>
                <td class="excel-cell number">${row.gemiddeldePerPartij.toFixed(2)}</td>
                <td class="excel-cell number">${row.totaalGemiddelde.toFixed(2)}</td>
                <td class="excel-cell number ${row.winstVerlies >= 0 ? 'percentage-positive' : 'percentage-negative'}">
                    ${row.winstVerlies.toFixed(1)}%
                </td>
                <td class="excel-cell number">${row.percHoogsteReeks.toFixed(1)}%</td>
                <td class="excel-cell number percentage-positive">${row.percGewonnen.toFixed(1)}%</td>
                <td class="excel-cell number percentage-negative">${row.percVerloren.toFixed(1)}%</td>
                <td class="excel-cell number">${row.matchpunten}</td>
            </tr>
        `;
    });
    
    html += `
                    </tbody>
                    
                    <tfoot>
                        <tr class="excel-summary-1">
                            <td colspan="3" class="summary-label">Totaal</td>
                            <td class="summary-value">${totaalPunten}</td>
                            <td class="summary-value">${totaalBeurten}</td>
                            <td class="summary-value">${maxHoogsteReeks}</td>
                            <td class="summary-value">${lastRow.aantalPartijen}</td>
                            <td class="summary-dash">-</td>
                            <td class="summary-value">${lastRow.totaalGewonnen}</td>
                            <td class="summary-dash">-</td>
                            <td class="summary-value">${lastRow.totaalGemiddelde.toFixed(2)}</td>
                            <td class="summary-value">${lastRow.winstVerlies.toFixed(1)}%</td>
                            <td class="summary-value">${lastRow.percHoogsteReeks.toFixed(1)}%</td>
                            <td class="summary-value">${lastRow.percGewonnen.toFixed(1)}%</td>
                            <td class="summary-value">${lastRow.percVerloren.toFixed(1)}%</td>
                            <td class="summary-value">${totaalMatchpunten}</td>
                        </tr>
                        
                        <tr class="excel-summary-2">
                            <td colspan="10" class="target-label">
                                Te maken punten volgend seizoen
                            </td>
                            <td class="target-calc">${(lastRow.totaalGemiddelde * 20).toFixed(2)}</td>
                            <td class="target-final">${calculateTargetValue(lastRow.totaalGemiddelde * 20, player.target)}</td>
                            <td colspan="4" class="warning-text">
                                ${lastRow.aantalPartijen < 20 ? '⚠️ Speler heeft geen 20 matchen gespeeld' : ''}
                            </td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    `;
    
    return html;
}

function calculateTargetValue(kValue, target) {
    if (!target || target === 0) return 20;
    
    let resultaatL = 20;
    
    if (kValue < 20) {
        resultaatL = 20;
    } else if (kValue > target) {
        resultaatL = kValue;
    } else if (kValue < (target * 0.9)) {
        resultaatL = target * 0.9;
    } else {
        resultaatL = kValue;
    }
    
    return Math.ceil(resultaatL);
}

function showSpelerDetail(spelerNaam) {
    console.log(`Laad Excel-stijl detail voor: ${spelerNaam}`);
    
    const players = getPlayers();
    const player = players.find(p => p.name === spelerNaam);
    
    if (!player) {
        document.getElementById('statsContent').innerHTML = `
            <div class="error-message">
                <h3>Speler niet gevonden</h3>
                <p>Speler "${spelerNaam}" bestaat niet in de database.</p>
            </div>
        `;
        return;
    }
    
    const excelTable = generateExcelStylePlayerDetail(spelerNaam);
    
    document.getElementById('statsContent').innerHTML = `
        <div class="excel-detail-view">
            <div class="detail-header">
                <button class="back-btn" onclick="loadStatsPage()">← Terug naar overzicht</button>
                <h2>${spelerNaam} - Excel Stijl Overzicht</h2>
                <div class="ref-values">
                    <div class="ref-item">
                        <span class="ref-label">TSG:</span>
                        <span class="ref-value">${player.tsg || 'N/A'}</span>
                    </div>
                    <div class="ref-item">
                        <span class="ref-label">Target:</span>
                        <span class="ref-value">${player.target || 'N/A'}</span>
                    </div>
                    <div class="ref-item">
                        <span class="ref-label">TSG eerste 6:</span>
                        <span class="ref-value">${player.tsgEersteZes || 'N/A'}</span>
                    </div>
                    <div class="ref-item">
                        <span class="ref-label">Target eerste 6:</span>
                        <span class="ref-value">${player.targetEersteZes || 'N/A'}</span>
                    </div>
                </div>
            </div>
            ${excelTable}
        </div>
    `;
}

// ==================== EXPORT FUNCTIES ====================
function exportMacroKlassement() {
    const klassement = berekenExactMacroKlassement();
    const { speeldagen, spelersKlassement } = klassement;
    const seizoen = getSeizoen();
    
    if (spelersKlassement.length === 0) {
        alert('Geen data om te exporteren');
        return;
    }
    
    let csv = '\uFEFF';
    csv += `"Klassement Biljartclub SPORT & VERMAAK ${seizoen}","","","","","","","","",""\n\n`;
    csv += '"Naam",';
    
    speeldagen.forEach((speeldag, index) => {
        if (index < speeldagen.length - 1) {
            csv += `"${speeldag.speeldagNummer}","","totaal",`;
        } else {
            csv += `"${speeldag.speeldagNummer}","","totaal","totaal"\n`;
        }
    });
    
    csv += '"",';
    speeldagen.forEach((speeldag, index) => {
        csv += `"${speeldag.datumShort}","",`;
    });
    csv += '\n';
    
    spelersKlassement.forEach(speler => {
        csv += `"${speler.positie}. ${speler.naam}",`;
        
        speeldagen.forEach((speeldag, dagIndex) => {
            const punten = speler.puntenPerDatum[speeldag.datum] || [];
            const match1 = punten[0] || '';
            const match2 = punten[1] || '';
            const tussentotaal = speler.tussentotalen[dagIndex] || 0;
            
            csv += `"${match1}","${match2}","${tussentotaal}",`;
        });
        
        csv += `"${speler.eindtotaal}"\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Klassement_Excel_Macro_${seizoen.replace('-', '_')}_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    alert(`Excel Macro Klassement geëxporteerd!\n\n${spelersKlassement.length} spelers\n${speeldagen.length} speeldagen\nSeizoen: ${seizoen}`);
}

function exportSpelerData(spelerNaam) {
    const matches = getMatches();
    const spelerMatches = matches.filter(m => 
        m.p1 === spelerNaam || m.p2 === spelerNaam
    );
    
    if (spelerMatches.length === 0) {
        alert('Geen data om te exporteren');
        return;
    }
    
    let csv = 'Datum;Tegenstander;Eigen Score;Tegenstander Score;Gewonnen;Hoogste Reeks;Matchpunten\n';
    
    spelerMatches.forEach(match => {
        const isPlayer1 = match.p1 === spelerNaam;
        const tegenstander = isPlayer1 ? match.p2 : match.p1;
        const eigenScore = isPlayer1 ? (match.p1Score || 0) : (match.p2Score || 0);
        const tegenScore = isPlayer1 ? (match.p2Score || 0) : (match.p1Score || 0);
        const gewonnen = match.winner === spelerNaam ? 'Ja' : 'Nee';
        const hoogsteReeks = isPlayer1 ? (match.p1Highest || 0) : (match.p2Highest || 0);
        const matchpunten = gewonnen === 'Ja' ? 2 : 1;
        
        csv += `${match.date};${tegenstander};${eigenScore};${tegenScore};${gewonnen};${hoogsteReeks};${matchpunten}\n`;
    });
    
    const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${spelerNaam}_matches_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    alert(`Match data van ${spelerNaam} geëxporteerd!`);
}

function exportPlayerExcel(spelerNaam) {
    exportSpelerData(spelerNaam);
}

function printMacroKlassement() {
    const klassementElement = document.querySelector('.excel-macro-klassement');
    if (!klassementElement) {
        alert('Klassement niet gevonden');
        return;
    }
    
    const printContent = klassementElement.innerHTML;
    const originalContent = document.body.innerHTML;
    
    document.body.innerHTML = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Klassement Biljart - Excel VBA Macro Stijl</title>
            <style>
                @page {
                    size: landscape;
                    margin: 0.5cm;
                }
                body {
                    font-family: 'Calibri', Arial, sans-serif;
                    margin: 0;
                    padding: 10px;
                    background: white;
                    color: black;
                }
                table {
                    border-collapse: collapse;
                    width: 100%;
                    font-size: 10pt;
                }
                td, th {
                    border: 1px solid #000;
                    padding: 2px 4px;
                    text-align: center;
                    height: 18px;
                }
                .title-cell {
                    background: #4472c4;
                    color: #c00000;
                    font-family: 'Comic Sans MS';
                    font-size: 20pt;
                    text-align: center;
                    height: 37.5px;
                    border: 1px solid #000;
                }
                .no-print {
                    display: none !important;
                }
                .print-info {
                    text-align: right;
                    font-size: 9pt;
                    color: #666;
                    margin-bottom: 10px;
                }
            </style>
        </head>
        <body>
            <div class="print-info">
                Geprint op: ${new Date().toLocaleDateString('nl-NL')} | 
                Klassement Biljartclub SPORT & VERMAAK ${getSeizoen()}
            </div>
            ${printContent}
        </body>
        </html>
    `;
    
    document.querySelectorAll('button, .no-print').forEach(el => el.style.display = 'none');
    
    window.print();
    
    document.body.innerHTML = originalContent;
    location.reload();
}

// ==================== HELPER FUNCTIES ====================
function showNoDataMessage(elementId, dataType) {
    const content = document.getElementById(elementId);
    if (content) {
        content.innerHTML = `
            <div class="no-data-message">
                <h3>Geen ${dataType} gevonden</h3>
                <p>Voeg eerst spelers en voltooide matches toe in de hoofdapp.</p>
                <button onclick="showPage(1)" class="action-btn">
                    ← Terug naar Home
                </button>
            </div>
        `;
    }
}

function showErrorMessage(elementId, error) {
    const content = document.getElementById(elementId);
    if (content) {
        content.innerHTML = `
            <div class="error-message">
                <h3>❌ Fout opgetreden</h3>
                <p><strong>${error.message}</strong></p>
                <button onclick="location.reload()" class="action-btn">
                    Herlaad Pagina
                </button>
            </div>
        `;
    }
}

function refreshKlassement() {
    if (confirm('Klassement herberekenen met macro-logica?\n\nAlle berekeningen worden opnieuw uitgevoerd zoals in de Excel VBA macro.')) {
        loadRankingPage();
    }
}

// ==================== GLOBAL EXPORTS ====================
window.loadStatsPage = loadStatsPage;
window.loadRankingPage = loadRankingPage;
window.showSpelerDetail = showSpelerDetail;
window.generateExcelStylePlayerDetail = generateExcelStylePlayerDetail;
window.createExcelTableHTML = createExcelTableHTML;
window.calculateTargetValue = calculateTargetValue;
window.exportMacroKlassement = exportMacroKlassement;
window.exportSpelerData = exportSpelerData;
window.exportPlayerExcel = exportPlayerExcel;
window.printMacroKlassement = printMacroKlassement;
window.refreshKlassement = refreshKlassement;

console.log('🎯 Excel VBA Macro stats.js functies geregistreerd');
