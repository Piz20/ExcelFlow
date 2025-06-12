// Fichier : Services/EmailContentBuilder.cs
using ExcelFlow.Utilities; // Pour accéder à GlobalConstants
using ExcelFlow.Models;   // <<< ADD THIS LINE for the EmailData class
using System; // Pour StringComparison

namespace ExcelFlow.Services;

public class EmailContentBuilder
{
    public string BuildSubject(EmailData emailData)
    {
        return GlobalConstants.SUBJECT_TEMPLATE
            .Replace("[NOM_PARTENAIRE]", emailData.PartnerNameInFile)
            .Replace("[DATE_FICHIER]", emailData.DateString);
    }

    public string BuildBody(EmailData emailData)
    {
        string dateDescriptor = "la journée du";
        // Cette logique est basée sur les formats que vous avez mentionnés.
        // Vous devrez peut-être l'ajuster si d'autres formats de date sont possibles.
        if (emailData.DateString.Contains("-") ||
            emailData.DateString.Contains(" au ", StringComparison.OrdinalIgnoreCase) ||
            emailData.DateString.Contains(","))
        {
            dateDescriptor = "les journées du";
        }

        return GlobalConstants.BODY_TEMPLATE
            .Replace("[JOURNEE_OU_JOURNEES_DU]", dateDescriptor)
            .Replace("[DATE_OU_INTERVALLE_JOURS_ANALYSE]", emailData.DateString)
            .Replace("[SOLDE_FINAL]", emailData.FinalBalance);
    }
}