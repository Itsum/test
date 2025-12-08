using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace DGP.CREX.Massenversand.Plugins.Accounts
{
    // Hilfsklasse für JSON
    [DataContract]
    public class EmailInput
    {
        [DataMember]
        public string TemplateId { get; set; }

        [DataMember]
        public string SenderType { get; set; }

        [DataMember]
        public string SenderId { get; set; }
    }

    public class Massenversand : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                tracer.Trace($"[{DateTime.Now:HH:mm:ss}] Plugin Massenversand gestartet. Message: {context.MessageName}");

                MassenversandService _massenVersandService = new MassenversandService(service, tracer);

                // --- SICHERHEITS-CHECK ---
                // Prüft, ob der User die Rolle "DGP Massenversand Manager" hat.
                // Falls nicht, wirft die Methode eine Exception und bricht hier ab.
                _massenVersandService.CheckUserHasRole(context.InitiatingUserId, "DGP Massenversand Manager");
                // -------------------------

                if (!context.InputParameters.Contains("Type") || !context.InputParameters.Contains("Input"))
                {
                    throw new InvalidPluginExecutionException("Parameter 'Type' oder 'Input' fehlen.");
                }

                string type = context.InputParameters["Type"].ToString();
                string inputRaw = context.InputParameters["Input"].ToString();

                tracer.Trace($"Action-Typ: '{type}'");

                if (type == "EmailVersenden")
                {
                    tracer.Trace("Zweig: E-Mail Versand");

                    EmailInput emailConfig = ParseJsonInput(inputRaw);

                    if (string.IsNullOrEmpty(emailConfig.TemplateId) || string.IsNullOrEmpty(emailConfig.SenderId))
                    {
                        throw new InvalidPluginExecutionException("JSON unvollständig (TemplateId/SenderId fehlen).");
                    }

                    EntityReference senderRef = _massenVersandService.ResolveSender(emailConfig.SenderType, new Guid(emailConfig.SenderId));
                    tracer.Trace($"Verwende Absender: {senderRef.LogicalName} ({senderRef.Id})");

                    List<Account> _accounts = _massenVersandService.RetrieveAccountsWhereVersandaktionEqualsTrue();
                    tracer.Trace($"{_accounts.Count} Empfänger gefunden.");

                    string result = _massenVersandService.CreateAndSendEmails(_accounts, emailConfig.TemplateId, senderRef);

                    context.OutputParameters["Output"] = result;
                }
                else if (type == "FirmenUpdate")
                {
                    tracer.Trace("Zweig: Firmen Update (CSV)");

                    List<Account> _accounts = _massenVersandService.ExtractAccountsFromMassenversandCSV(inputRaw);
                    string result = _massenVersandService.BulkUpdateAccounts(_accounts);
                    context.OutputParameters["Output"] = result;
                }
                else
                {
                    throw new InvalidPluginExecutionException($"Unbekannter Typ: {type}");
                }

                tracer.Trace("Plugin erfolgreich beendet.");
            }
            catch (Exception ex)
            {
                tracer.Trace($"FEHLER: {ex.Message}\n{ex.StackTrace}");
                throw new InvalidPluginExecutionException($"Fehler: {ex.Message}", ex);
            }
        }

        private EmailInput ParseJsonInput(string json)
        {
            try
            {
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    var serializer = new DataContractJsonSerializer(typeof(EmailInput));
                    return (EmailInput)serializer.ReadObject(ms);
                }
            }
            catch (Exception)
            {
                return new EmailInput { TemplateId = json, SenderType = null, SenderId = null };
            }
        }
    }
}
