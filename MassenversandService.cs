using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DGP.CREX.Massenversand.Plugins.Accounts
{
    public class MassenversandService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _tracer;

        public MassenversandService(IOrganizationService service, ITracingService tracer)
        {
            _service = service;
            _tracer = tracer;
        }

        /// <summary>
        /// NEU: Prüft, ob der Benutzer eine bestimmte Sicherheitsrolle hat.
        /// </summary>
        public void CheckUserHasRole(Guid userId, string roleName)
        {
            _tracer.Trace($"Prüfe Sicherheitsrolle '{roleName}' für User {userId}...");

            QueryExpression query = new QueryExpression("role");
            query.ColumnSet.AddColumn("name");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, roleName);

            // Verknüpfung: Role -> SystemUserRoles -> SystemUser
            // Wir suchen nur nach Rollen, die diesem User direkt zugewiesen sind.
            LinkEntity link = query.AddLink("systemuserroles", "roleid", "roleid");
            link.LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);

            EntityCollection results = _service.RetrieveMultiple(query);

            if (results.Entities.Count > 0)
            {
                _tracer.Trace($"Berechtigung bestätigt: User hat die Rolle '{roleName}'.");
            }
            else
            {
                _tracer.Trace("Berechtigung fehlt: Rolle nicht gefunden.");
                throw new InvalidPluginExecutionException($"Zugriff verweigert: Sie benötigen die Sicherheitsrolle '{roleName}', um diese Aktion auszuführen.");
            }
        }

        /// <summary>
        /// Ermittelt den korrekten Absender (SystemUser oder Queue) basierend auf der Eingabe.
        /// </summary>
        public EntityReference ResolveSender(string type, Guid id)
        {
            if (type.Equals("systemuser", StringComparison.OrdinalIgnoreCase))
            {
                return new EntityReference("systemuser", id);
            }
            else if (type.Equals("dgp_usergroup", StringComparison.OrdinalIgnoreCase))
            {
                _tracer.Trace($"Löse Queue für User Group '{id}' auf...");
                try
                {
                    Entity group = _service.Retrieve("dgp_usergroup", id, new ColumnSet("dgp_queueid"));
                    EntityReference queueRef = group.GetAttributeValue<EntityReference>("dgp_queueid");

                    if (queueRef == null)
                        throw new InvalidPluginExecutionException("Die ausgewählte Benutzergruppe hat keine Queue (dgp_queueid) hinterlegt.");

                    _tracer.Trace($"Queue gefunden: {queueRef.Name} ({queueRef.Id})");
                    return queueRef;
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"Fehler beim Auflösen der User Group Queue: {ex.Message}");
                }
            }
            else
            {
                throw new InvalidPluginExecutionException($"Ungültiger Absender-Typ: {type}");
            }
        }

        public List<Account> RetrieveAccountsWhereVersandaktionEqualsTrue()
        {
            _tracer.Trace("Starte Abruf der Accounts (Versandaktion = true)...");

            QueryExpression query = new QueryExpression(Account.EntityLogicalName);
            query.ColumnSet.AddColumns("dgp_mailingcampaign_logic", "name", "primarycontactid");
            query.Criteria.AddCondition("dgp_mailingcampaign_logic", ConditionOperator.Equal, true);

            EntityCollection result = _service.RetrieveMultiple(query);
            _tracer.Trace($"Abruf beendet. {result.Entities.Count} Datensätze gefunden.");

            return result.Entities.Select(e => e.ToEntity<Account>()).ToList();
        }

        public List<Account> ExtractAccountsFromMassenversandCSV(string configIdString)
        {
            _tracer.Trace($"Beginne CSV-Extraktion aus Konfiguration ID: {configIdString}");
            string csvContent = RetrieveAccountsFileFromConfiguration(configIdString);

            var result = new List<Account>();
            var lines = csvContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines.Skip(1))
            {
                var columns = line.Split(';');
                if (columns.Length < 3) continue;

                if (!Guid.TryParse(columns[0].Trim(), out Guid accId)) continue;

                bool versandAktion = columns[2].Trim().Equals("true", StringComparison.OrdinalIgnoreCase)
                                     || columns[2].Trim().Equals("ja", StringComparison.OrdinalIgnoreCase);

                var account = new Account
                {
                    Id = accId,
                    dgp_mailingcampaign_logic = versandAktion
                };
                result.Add(account);
            }
            return result;
        }

        public string BulkUpdateAccounts(List<Account> accounts)
        {
            if (accounts == null || accounts.Count == 0) return "Keine Daten.";

            EntityCollection accountsEC = new EntityCollection(accounts.Cast<Entity>().ToList())
            {
                EntityName = Account.EntityLogicalName
            };

            var updateRequest = new UpdateMultipleRequest { Targets = accountsEC };
            _service.Execute(updateRequest);

            return $"{accounts.Count} Firmen wurden erfolgreich aktualisiert.";
        }

        public string CreateAndSendEmails(List<Account> accounts, string templateId, EntityReference sender)
        {
            _tracer.Trace($"Starte E-Mail Prozess. TemplateId: {templateId}");

            int sentCount = 0;
            int skippedCount = 0;

            foreach (var account in accounts)
            {
                try
                {
                    if (account.PrimaryContactId == null)
                    {
                        skippedCount++;
                        continue;
                    }

                    // 1. Template instanziieren
                    var instantiateRequest = new InstantiateTemplateRequest
                    {
                        TemplateId = new Guid(templateId),
                        ObjectId = account.Id,
                        ObjectType = Account.EntityLogicalName
                    };

                    var response = (InstantiateTemplateResponse)_service.Execute(instantiateRequest);
                    if (response.EntityCollection.Entities.Count == 0) continue;

                    var email = response.EntityCollection.Entities[0];

                    // 2. Empfänger (To) setzen
                    email["to"] = new EntityCollection(new List<Entity>
                    {
                        new Entity("activityparty") { ["partyid"] = account.PrimaryContactId }
                    });

                    // 3. Absender (From) setzen
                    if (sender != null)
                    {
                        email["from"] = new EntityCollection(new List<Entity>
                        {
                            new Entity("activityparty") { ["partyid"] = sender }
                        });
                    }

                    // 4. Bezug setzen
                    email["regardingobjectid"] = account.ToEntityReference();

                    // 5. E-Mail erstellen & senden
                    Guid emailId = _service.Create(email);

                    var sendRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        IssueSend = true
                    };
                    _service.Execute(sendRequest);

                    sentCount++;
                }
                catch (Exception ex)
                {
                    _tracer.Trace($"Fehler bei Account {account.Id}: {ex.Message}");
                }
            }

            return $"Fertig. Gesendet: {sentCount}, Übersprungen: {skippedCount}.";
        }

        private string RetrieveAccountsFileFromConfiguration(string id)
        {
            Guid configId = new Guid(id);
            Entity configRecord = _service.Retrieve("dgp_configuration", configId, new ColumnSet("dgp_annotation", "dgp_annotation_name"));
            string fileName = configRecord.GetAttributeValue<string>("dgp_annotation_name");

            var initFileRequest = new InitializeFileBlocksDownloadRequest
            {
                Target = new EntityReference("dgp_configuration", configId),
                FileAttributeName = "dgp_annotation"
            };

            var initFileResponse = (InitializeFileBlocksDownloadResponse)_service.Execute(initFileRequest);

            var downloadRequest = new DownloadBlockRequest
            {
                FileContinuationToken = initFileResponse.FileContinuationToken,
                BlockLength = initFileResponse.FileSizeInBytes,
                Offset = 0
            };

            var downloadResponse = (DownloadBlockResponse)_service.Execute(downloadRequest);
            return Encoding.UTF8.GetString(downloadResponse.Data);
        }
    }
}
