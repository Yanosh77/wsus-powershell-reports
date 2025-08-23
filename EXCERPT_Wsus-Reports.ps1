# EXCERPT — The full script is available for purchase.
# This excerpt is intentionally non-functional and shows only structure & style.
<#
  Wsus-Reports_Universal_multilingual_email.ps1
  Rapport HTML WSUS multilingue (FR / EN / ES) avec auto-détection de langue.
  - Auto WSUS (Registre) + override -Server/-UseSSL/-Port
  - Module UpdateServices ou fallback DLL
  - JS ES5, perfs ExplainDays, CSV groupes, nettoyage MaxKeep
  - libellés HTML + messages console + parsing anciens rapports (FR/EN/ES)
#>

param(
  # --- Langue ---
  [string]$Lang = 'en',              # 'auto' | 'fr' | 'en' | 'es'

  # --- Connexion WSUS (laisser vides pour autodétection locale) ---
  [string]$Server,
  [Nullable[bool]]$UseSSL,
  [int]$Port,

  # --- Options globales ---
  [int]   $StaleDays = 30,
  [string]$OutDir    = 'C:\WSUS-Reports',
  [switch]$OpenWhenDone,
  [switch]$ExportCsv,
  [int]   $MaxKeep   = 3,

  # --- Onglet "Mises a jour recentes" ---
  [int]   $DaysRecent = 30,
  [int]   $MaxRecentUpdates = 30,
  [bool]  $ApprovedOnly = $true,
  [bool]  $IncludeSuperseded = $false,

  # --- Onglet "Pourquoi ces postes ?" ---
  [int]   $ExplainTop = 15,
  [bool]  $ExplainApprovedOnly = $false,
  [bool]  $ExplainIncludeSuperseded = $true,
  [int]   $ExplainDays = 120,

  # --- Onglet Upgrades ---
  [bool]  $AddUpgradesTab = $true,
  [bool]  $UpgradesApprovedOnly = $false,
  
  # --- Email (SMTP uniquement) ---
  [ValidateSet('None','SMTP')]
  [string]$MailMode = 'None',

  [string]$MailFrom,
  [string]$MailTo,
  [string]$MailCc,
  [string]$MailBcc,
  [string]$MailSubject,
  [string]$MailBodyHtml,

  # SMTP
  [string]$SmtpServer,
  [int]$SmtpPort = 25,
  [bool]$SmtpUseSsl = $false,
  [string]$SmtpUser,
  [securestring]$SmtpPassword
)


# --- BLOC CONFIG EMAIL (centralisé, SMTP relay) ---
$EmailConfig = @{
  Enabled       = $true                # passe à $false pour désactiver l'envoi
  MailMode      = 'SMTP'               # 'SMTP' ou 'None'
  MailFrom      = 'xxx@xxx.xx'
  MailTo        = 'xxx@xxx.xx'
  MailCc        = ''
  MailBcc       = ''
  MailSubject   = 'WSUS Report'                   # vide = objet auto (i18n)
  MailBodyHtml  = ''                   # vide = corps auto (i18n)

  SmtpServer    = 'xxx'
  SmtpPort      = 25
  SmtpUseSsl    = $true               # relay interne = false
  SmtpUser      = 'xxx@xxx.xx'                   # vide = anonyme
  SmtpPasswordPath = ''                # optionnel: chemin d'un SecureString chiffré (DPAPI)
}

# --- MERGE CONFIG -> VARIABLES si pas déjà passées en paramètre ---
if ($EmailConfig.Enabled) {
  if ($MailMode -eq 'None') { $MailMode = $EmailConfig.MailMode }
  foreach($k in 'MailFrom','MailTo','MailCc','MailBcc','MailSubject','MailBodyHtml','SmtpServer','SmtpPort','SmtpUseSsl','SmtpUser'){
    if(-not $PSBoundParameters.ContainsKey($k) -and $EmailConfig[$k] -ne $null){
      Set-Variable -Name $k -Value $EmailConfig[$k] -Scope Script
    }
  }
  if(-not $SmtpPassword -and $EmailConfig.SmtpPasswordPath -and (Test-Path $EmailConfig.SmtpPasswordPath)){
    $SmtpPassword = Get-Content $EmailConfig.SmtpPasswordPath | ConvertTo-SecureString
  }
}



# ---------- i18n ----------
function Resolve-Lang([string]$l){
  if([string]::IsNullOrWhiteSpace($l) -or $l -eq 'auto'){
    try { $l = [System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName } catch { $l = 'en' }
  }
  switch ($l.ToLower()){
    'fr' { 'fr' }
    'es' { 'es' }
    default { 'en' }
  }
}
$Lang = Resolve-Lang $Lang

function BoolStr([bool]$b){
  switch ($Lang){
    'fr' { if($b){'Oui'}else{'Non'} }
    'es' { if($b){'Sí'}else{'No'} }
    default { if($b){'Yes'}else{'No'} }
  }
}

# Libellés par langue
$L = @{
  'fr' = @{
    header_title = 'Rapport WSUS - Conformité des mises à jour'
    tab_overview = 'Aperçu'
    tab_updates  = 'Mises à jour récentes'
    tab_explain  = 'Top Maj Manquantes ?'
    tab_upgrades = 'Upgrades'
    kpi_machines = 'Postes'
    kpi_compliance = 'Taux de conformité'
    kpi_uptodate = 'À jour'
    kpi_need     = 'Nécessitent MAJ'
    kpi_errors   = 'En erreur'
    kpi_stale_fm = 'Inactifs > {0} j'
    kpi_unknown  = 'Inconnus'
    kpi_upg_uniq = 'Upgrades en attente (uniques)'
    evo_title_fm = 'Évolution (Δ vs {0} / {1})'
    evo_metric   = 'Métrique'
    evo_current  = 'Actuel'
    evo_vs_s1    = 'Δ vs S-1'
    evo_vs_s2    = 'Δ vs S-2'
    m_up         = 'À jour'
    m_need       = 'Nécessitent MAJ'
    m_err        = 'En erreur'
    m_stale      = 'Inactifs'
    m_machines   = 'Postes'
    m_compliance = 'Conformité (%)'
    groups_title = 'Par groupes WSUS'
    th_group='Groupe'; th_total='Total'; th_up='À jour'; th_need='Nécessitent'; th_err='Erreurs'; th_unk='Inconnu'; th_stale='Inactifs'; th_comp='Conformité'; th_vis='Visuel'
    upd_pill_fm  = 'Fenêtre: {0} jours | Max: {1} updates | Approuvées seulement: {2} | Inclure "remplacées": {3}'
    th_date='Date'; th_title='Titre'; th_need2='Nécessitent'; th_installed='Installées'; th_fail='Échecs'; th_details='Détails'
    state_dl='téléchargée'; state_pr='attente redémarrage'; state_ni='pas téléchargée'
    view_need_fm='Voir ({0}) — téléchargée ({1}), attente redémarrage ({2}), pas téléchargée ({3})'
    fail_fm='Échecs ({0})'
    explain_pill='Top des mises à jour manquantes. Comptage par mise à jour (un poste peut apparaître plusieurs fois).'
    explain_none='Aucune mise à jour avec des postes en "Nécessitent" n''a été trouvée avec les filtres actuels. Essayez ExplainApprovedOnly=false et ExplainIncludeSuperseded=true.'
    upg_pill='Mises à jour de fonctionnalité (Feature updates). Les compteurs sont dédupliqués par MAJ, et le KPI montre le nombre de machines uniques ayant au moins une upgrade.'
    upg_kpi='Machines avec upgrade (uniques)'
    upg_none='Aucune feature update en attente détectée.'
    footer_fm='Généré le {0} - Serveur {1}:{2} (SSL={3})'
    c_connect='Connexion à WSUS {0}:{1} (SSL={2})...'
    c_groups='Calcul des états par groupe...'
    c_recent='Collecte des mises à jour récentes...'
    c_analyze='Analyse détaillée de {0} updates...'
    c_explain='Construction de l''onglet explicatif (Top {0} MAJ)...'
    c_written='Rapport écrit : {0}'
    c_csv='Export CSV : {0}'
  }
  'es' = @{
    header_title = 'Informe de WSUS - Cumplimiento de actualizaciones'
    tab_overview = 'Resumen'
    tab_updates  = 'Actualizaciones recientes'
    tab_explain  = '¿Top Actualizaciones faltadas?'
    tab_upgrades = 'Actualizaciones de funciones'
    kpi_machines = 'Equipos'
    kpi_compliance = 'Tasa de cumplimiento'
    kpi_uptodate = 'Actualizados'
    kpi_need     = 'Requieren act.'
    kpi_errors   = 'Con errores'
    kpi_stale_fm = 'Inactivos > {0} d'
    kpi_unknown  = 'Desconocidos'
    kpi_upg_uniq = 'Actualizaciones pendientes (únicas)'
    evo_title_fm = 'Evolución (Δ vs {0} / {1})'
    evo_metric   = 'Métrica'
    evo_current  = 'Actual'
    evo_vs_s1    = 'Δ vs S-1'
    evo_vs_s2    = 'Δ vs S-2'
    m_up         = 'Actualizados'
    m_need       = 'Requieren act.'
    m_err        = 'Con errores'
    m_stale      = 'Inactivos'
    m_machines   = 'Equipos'
    m_compliance = 'Cumplimiento (%)'
    groups_title = 'Por grupos WSUS'
    th_group='Grupo'; th_total='Total'; th_up='Actualizados'; th_need='Requieren'; th_err='Errores'; th_unk='Desconocido'; th_stale='Inactivos'; th_comp='Cumplimiento'; th_vis='Visual'
    upd_pill_fm  = 'Ventana: {0} días | Máx: {1} act. | Solo aprobadas: {2} | Incluir reemplazadas: {3}'
    th_date='Fecha'; th_title='Título'; th_need2='Requieren'; th_installed='Instaladas'; th_fail='Fallos'; th_details='Detalles'
    state_dl='descargada'; state_pr='pendiente reinicio'; state_ni='no descargada'
    view_need_fm='Ver ({0}) — descargada ({1}), pendiente reinicio ({2}), no descargada ({3})'
    fail_fm='Fallos ({0})'
    explain_pill='Top actualizaciones faltadas". Conteo por actualización (un equipo puede aparecer varias veces).'
    explain_none='No se encontraron actualizaciones con equipos en "Requieren act." con los filtros actuales. Prueba ExplainApprovedOnly=false y ExplainIncludeSuperseded=true.'
    upg_pill='Actualizaciones de características (Feature updates). Contadores deduplicados por actualización; el KPI muestra el número de equipos únicos con al menos una actualización.'
    upg_kpi='Equipos con actualización (únicos)'
    upg_none='No se detectaron actualizaciones de características pendientes.'
    footer_fm='Generado el {0} - Servidor {1}:{2} (SSL={3})'
    c_connect='Conexión a WSUS {0}:{1} (SSL={2})...'
    c_groups='Cálculo de estados por grupo...'
    c_recent='Recopilando actualizaciones recientes...'
    c_analyze='Análisis detallado de {0} actualizaciones...'
    c_explain='Construyendo la pestaña de explicación (Top {0})...'
    c_written='Informe escrito: {0}'
    c_csv='Exportación CSV: {0}'
  }
  'en' = @{
    header_title = 'WSUS Report - Update Compliance'
    tab_overview = 'Overview'
    tab_updates  = 'Recent updates'
    tab_explain  = 'Top missing Updates?'
    tab_upgrades = 'Upgrades'
    kpi_machines = 'Machines'
    kpi_compliance = 'Compliance rate'
    kpi_uptodate = 'Up-to-date'
    kpi_need     = 'Need updates'
    kpi_errors   = 'Errors'
    kpi_stale_fm = 'Stale > {0} d'
    kpi_unknown  = 'Unknown'
    kpi_upg_uniq = 'Unique pending upgrades'
    evo_title_fm = 'Evolution (Δ vs {0} / {1})'
    evo_metric   = 'Metric'
    evo_current  = 'Current'
    evo_vs_s1    = 'Δ vs S-1'
    evo_vs_s2    = 'Δ vs S-2'
    m_up         = 'Up-to-date'
    m_need       = 'Need updates'
    m_err        = 'Errors'
    m_stale      = 'Stale'
    m_machines   = 'Machines'
    m_compliance = 'Compliance (%)'
    groups_title = 'By WSUS groups'
    th_group='Group'; th_total='Total'; th_up='Up-to-date'; th_need='Need'; th_err='Errors'; th_unk='Unknown'; th_stale='Stale'; th_comp='Compliance'; th_vis='Visual'
    upd_pill_fm  = 'Window: {0} days | Max: {1} updates | Approved only: {2} | Include superseded: {3}'
    th_date='Date'; th_title='Title'; th_need2='Need'; th_installed='Installed'; th_fail='Failures'; th_details='Details'
    state_dl='downloaded'; state_pr='pending reboot'; state_ni='not downloaded'
    view_need_fm='View ({0}) — downloaded ({1}), pending reboot ({2}), not downloaded ({3})'
    fail_fm='Failures ({0})'
    explain_pill='Top missing updates. Counted per update (a machine may appear multiple times).'
    explain_none='No update with machines in "Need" found with current filters. Try ExplainApprovedOnly=false and ExplainIncludeSuperseded=true.'
    upg_pill='Feature updates. Counters are deduped per update; KPI shows unique machines with at least one upgrade.'
    upg_kpi='Machines with upgrade (unique)'
    upg_none='No feature update pending detected.'
    footer_fm='Generated on {0} - Server {1}:{2} (SSL={3})'
    c_connect='Connecting to WSUS {0}:{1} (SSL={2})...'
    c_groups='Calculating group states...'
    c_recent='Collecting recent updates...'
    c_analyze='Detailed analysis of {0} updates...'
    c_explain='Building explanation tab (Top {0})...'
    c_written='Report written: {0}'
    c_csv='CSV export: {0}'
  }
}[$Lang]