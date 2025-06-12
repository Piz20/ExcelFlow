// Fichier : Utilities/GlobalConstants.cs

namespace ExcelFlow.Utilities;

public static class GlobalConstants
{
    public const string SUBJECT_TEMPLATE = "Objet : Détermination du solde final de [NOM_PARTENAIRE] - [DATE_FICHIER]";

    public const string BODY_TEMPLATE = @"
<p>Bonsoir cher partenaire,</p>
<p>Merci de trouver en pièce jointe ci-dessous l'analyse de votre compte support pour [JOURNEE_OU_JOURNEES_DU] <strong>[DATE_OU_INTERVALLE_JOURS_ANALYSE]</strong>, ayant permis d'aboutir au solde final (avant cash out auto) de <span style='color: #f51b1b;'><strong>[SOLDE_FINAL] [CURRENCY]</strong></span>.</p>
<p>Vous trouverez également les éléments ayant servi de base au calcul de votre solde final, pour [JOURNEE_OU_JOURNEES_DU] <strong>[DATE_OU_INTERVALLE_JOURS_ANALYSE]</strong>.</p>
<p>Nous restons disponibles pour toute information complémentaire.</p>
<p><span style='color: #f51b1b;'><strong>NB :</strong> Prière d'effectuer vos contrôles caisse et compte support à J+1 (dans les 24H) afin de nous remonter toute anomalie constatée pour régularisation, soit à notre niveau, soit au niveau du MTO.</span></p><p>Merci d'accuser réception.</p>
<p>Cordialement,</p>";
}