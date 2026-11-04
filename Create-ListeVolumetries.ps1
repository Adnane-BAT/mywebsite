# Crée la liste SharePoint "Volumetries" + données de la maquette (CEPAC – Annuelle 2026)
# Prérequis : Install-Module PnP.PowerShell -Scope CurrentUser
$SiteUrl  = "https://<tenant>.sharepoint.com/sites/<site>"
$ClientId = "<app-registration-client-id>"   # requis par PnP.PowerShell 2.x+
$List     = "Volumetries"

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

New-PnPList -Title $List -Template GenericList -OnQuickLaunch | Out-Null
Set-PnPField -List $List -Identity "Title" -Values @{ Title = "Segment" }

Add-PnPField -List $List -DisplayName "Entite"                   -InternalName "Entite"                   -Type Text     -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Campagne"                 -InternalName "Campagne"                 -Type Text     -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Niveau"                   -InternalName "Niveau"                   -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Ordre"                    -InternalName "Ordre"                    -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "VolumeAnnuel"             -InternalName "VolumeAnnuel"             -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Encours"                  -InternalName "Encours"                  -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "TauxDefaut"               -InternalName "TauxDefaut"               -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "VolumeEchantillonPropose" -InternalName "VolumeEchantillonPropose" -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "VolumeDemande"            -InternalName "VolumeDemande"            -Type Number   -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Commentaire"              -InternalName "Commentaire"              -Type Note     -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "Statut"                   -InternalName "Statut"                   -Type Text     -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "ValidePar"                -InternalName "ValidePar"                -Type Text     -AddToDefaultView | Out-Null
Add-PnPField -List $List -DisplayName "DateValidation"           -InternalName "DateValidation"           -Type DateTime -AddToDefaultView | Out-Null

# Encours en M€, TauxDefaut en % (5.0 = 5,0 %)
$Rows = @(
  @{ N=1; O=1; S="Mettre en place un crédit MLT Pro corp";  V=5385;  E=805.8; T=5.0; P=80 },
  @{ N=1; O=2; S="Mettre en place un crédit MLT Corp LF";   V=13535; E=456.4; T=4.7; P=80 },
  @{ N=1; O=3; S="LBO";                                     V=8393;  E=184.1; T=1.5; P=80 },
  @{ N=1; O=4; S="Provision S3 Retail Pro - N1";            V=4023;  E=78.9;  T=5.4; P=90 },
  @{ N=1; O=5; S="Forbearance Retail Professionnel";        V=2108;  E=805.4; T=7.8; P=40 },
  @{ N=1; O=6; S="Forbearance Corporate";                   V=1093;  E=491.4; T=5.1; P=35 },
  @{ N=2; O=1; S="Crédit à la consommation";                V=843;   E=474.1; T=2.0; P=90 }
)

foreach ($r in $Rows) {
  Add-PnPListItem -List $List -Values @{
    Title                    = $r.S
    Entite                   = "CEPAC"
    Campagne                 = "Annuelle 2026"
    Niveau                   = $r.N
    Ordre                    = $r.O
    VolumeAnnuel             = $r.V
    Encours                  = $r.E
    TauxDefaut               = $r.T
    VolumeEchantillonPropose = $r.P
    Statut                   = "À valider"
  } | Out-Null
}
Write-Host "Liste '$List' créée avec $($Rows.Count) segments." -ForegroundColor Green
