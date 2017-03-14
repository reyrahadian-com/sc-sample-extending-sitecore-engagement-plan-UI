using System;
using System.Collections.Generic;
using System.Web;
using Sitecore;
using Sitecore.Analytics;
using Sitecore.Analytics.Automation.MarketingAutomation;
using Sitecore.Analytics.Data;
using Sitecore.Analytics.DataAccess;
using Sitecore.Analytics.Model;
using Sitecore.Analytics.Model.Entities;
using Sitecore.Analytics.Tracking;
using Sitecore.Configuration;
using Sitecore.Data;
using Sitecore.Diagnostics;
using Sitecore.Globalization;
using Sitecore.IO;
using Sitecore.Jobs;
using Sitecore.Resources;
using Sitecore.Web.UI.HtmlControls;
using Sitecore.Web.UI.Sheer;

namespace TC.SitecoreEngagementPlanUI.Shell.Applications.MarketingAutomations.Dialogs
{
    public class ImportVisitorsForm : Sitecore.Shell.Applications.MarketingAutomation.Dialogs.ImportVisitorsForm
    {
        protected EditableCombobox ContactEmail;
        protected EditableCombobox ContactFirstName;
        protected EditableCombobox ContactIdentifier;
        protected EditableCombobox ContactLastName;
        protected Radiobutton ImportContacts;
        protected Radiobutton ImportUsers;
        protected Literal LiteralBadUserName;
        protected Literal LiteralBrokenStructure;
        protected Literal LiteralUserExists;
        protected Literal LiteralUsersImported;

        protected ImportActionType ImportAction
        {
            get
            {
                if (ImportUsers.Checked)
                    return ImportActionType.ImportAsUser;

                if (ImportContacts.Checked)
                    return ImportActionType.ImportAsContact;

                return ImportActionType.Unknown;
            }
        }

        protected override bool ActivePageChanging(string page, ref string newpage)
        {
            Assert.ArgumentNotNull(page, "page");
            Assert.ArgumentNotNull(newpage, "newpage");
            switch (page)
            {
                case "ChooseAction":
                    if (ImportAction == ImportActionType.Unknown)
                    {
                        SheerResponse.Alert("Please select one of the action available");
                        return false;
                    }
                    break;
                case "SelectFile":
                    if (newpage.Equals("Fields") && !ReadFile())
                        return false;
                    if ((ImportAction == ImportActionType.ImportAsContact) && newpage.Equals("Fields"))
                        newpage = "ContactFields";
                    break;
                case "Fields":
                    if (newpage.Equals("DomainAndRole") && !CheckFields())
                        return false;
                    break;
                case "DomainAndRole":
                    if (newpage.Equals("ContactFields"))
                    {
                        if (!CheckDomainAndRole())
                            return false;
                        StartImport();
                        newpage = "Importing";
                    }
                    break;

                case "ContactFields":
                    if (newpage.Equals("Importing"))
                    {
                        if (ImportAction == ImportActionType.ImportAsContact)
                        {
                            if (ContactIdentifier.Value.Equals(FormatDefaultText(Translate.Text("select field"))))
                            {
                                SheerResponse.Alert("Contact Identifier field must have a valid value");
                                return false;
                            }
                            StartImportingContacts();
                        }
                    }
                    else if (newpage.Equals("DomainAndRole"))
                    {
                        if (ImportAction == ImportActionType.ImportAsContact)
                            newpage = "SelectFile";
                    }
                    break;
                case "Importing":
                    if (newpage.Equals("Finish"))
                    {
                        if (ImportAction == ImportActionType.ImportAsContact)
                        {
                            LiteralUserExists.Text = "Contact exists:";
                            LiteralBadUserName.Text = "Invalid contact:";
                            LiteralUsersImported.Text = "Contacts imported:";
                            LiteralBrokenStructure.Visible = false;
                            NumBroken.Visible = false;
                        }
                        if (ImportAction == ImportActionType.ImportAsUser)
                        {
                            LiteralUserExists.Text = "User exists:";
                            LiteralBadUserName.Text = "Invalid user name:";
                            LiteralUsersImported.Text = "Users Imported:";
                            LiteralBrokenStructure.Visible = true;
                            NumBroken.Visible = true;
                        }
                    }
                    break;
            }
            return base.ActivePageChanging(page, ref newpage);
        }

        public new void Next()
        {
            var index = Pages.IndexOf(Active) + 1;
            if (index >= Pages.Count)
                Active = Pages[Pages.Count - 1];
            else
                Active = Pages[index];
        }

        private void InitializeFields()
        {
            for (var index = ParentControl.Controls.Count - 1; index >= 0; --index)
            {
                var control = ParentControl.Controls[index];
                if (!string.IsNullOrEmpty(control.ID) &&
                    (control.ID.StartsWith("Field") || control.ID.StartsWith("Property") ||
                     control.ID.StartsWith("DelSection")))
                    ParentControl.Controls.Remove(control);
            }
            InitializeUserNameField();
            InitializeContactFields();
            var prevIdPostfix = string.Empty;
            foreach (var str in FieldList)
            {
                var f = str;
                var defProperty = PropertyList.Find(p => string.Equals(f, p, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(defProperty))
                    prevIdPostfix = AddRow(prevIdPostfix, defProperty);
            }
            AddRow(prevIdPostfix);
            SheerResponse.Refresh(FieldsSection);
        }

        private void InitializeContactFields()
        {
            ContactIdentifier.List = FieldList;
            ContactIdentifier.Value = FormatDefaultText(Translate.Text("select field"));
            ContactFirstName.List = FieldList;
            ContactFirstName.Value = FormatDefaultText(Translate.Text("select field"));
            ContactLastName.List = FieldList;
            ContactLastName.Value = FormatDefaultText(Translate.Text("select field"));
            ContactEmail.List = FieldList;
            ContactEmail.Value = FormatDefaultText(Translate.Text("select field"));
        }

        protected void StartImportingContacts()
        {
            var importOptions = new ContactImportOptions
            {
                FileName = RealFilename,
                PlanId = Tracker.DefinitionDatabase.GetItem(StateId).ParentID,
                StateId = StateId,
                ContactIdentifierField = ContactIdentifier.Value,
                ContactFirstNameField = ContactFirstName.Value,
                ContactLastNameField = ContactLastName.Value,
                ContactEmailField = ContactEmail.Value
            };
            StartJob("Import Contacts", "DoImportContacts", this, importOptions);
            CheckImport();
        }

        protected string DoImportContacts(ContactImportOptions options)
        {
            Assert.ArgumentNotNull(options, "options");
            var contactManager = Factory.CreateObject("tracking/contactManager", true) as ContactManager;
            var contactRepository = Factory.CreateObject("tracking/contactRepository", true) as ContactRepository;
            var numOfImported = 0;
            var numOfContactExists = 0;
            var numOfBadContacts = 0;
            var numOfContactsAddedToState = 0;
            using (var csvFileReader = new CsvFileReader(options.FileName))
            {
                try
                {
                    var columnNameFields = csvFileReader.ReadLine();
                    var contactIdentifierIndex =
                        columnNameFields.FindIndex(h => string.Equals(options.ContactIdentifierField, h));
                    if (contactIdentifierIndex < 0)
                        return string.Empty;

                    var contactFirstNameIndex =
                        columnNameFields.FindIndex(x => string.Equals(options.ContactFirstNameField, x));
                    var contactLastNameIndex =
                        columnNameFields.FindIndex(x => string.Equals(options.ContactLastNameField, x));
                    var contactEmailIndex = columnNameFields.FindIndex(x => string.Equals(options.ContactEmailField, x));

                    var valueFields = csvFileReader.ReadLine();
                    while (valueFields != null)
                    {
                        var contactIdentifier = valueFields[contactIdentifierIndex];
                        if (string.IsNullOrWhiteSpace(contactIdentifier))
                        {
                            numOfBadContacts++;
                        }
                        else
                        {
                            var leaseOwner = new LeaseOwner("AddContacts-" + Guid.NewGuid(),
                                LeaseOwnerType.OutOfRequestWorker);
                            LockAttemptResult<Contact> lockAttemptResult;
                            var contact = contactManager.LoadContactReadOnly(contactIdentifier);
                            if (contact != null)
                            {
                                numOfContactExists++;
                                lockAttemptResult = contactRepository.TryLoadContact(contact.ContactId, leaseOwner,
                                    TimeSpan.FromSeconds(3));
                                if (lockAttemptResult.Status == LockAttemptStatus.Success)
                                {
                                    contact = lockAttemptResult.Object;
                                    contact.ContactSaveMode = ContactSaveMode.AlwaysSave;
                                }
                                else
                                {
                                    Log.Error("Cannot lock contact! " + lockAttemptResult.Status, this);
                                }
                            }
                            else
                            {
                                contact = contactRepository.CreateContact(ID.NewID);
                                contact.Identifiers.Identifier = contactIdentifier;
                                contact.System.Value = 0;
                                contact.System.VisitCount = 0;
                                contact.ContactSaveMode = ContactSaveMode.AlwaysSave;
                                lockAttemptResult = new LockAttemptResult<Contact>(LockAttemptStatus.Success, contact,
                                    leaseOwner);
                            }
                            UpdateContactPersonalInfo(contact, contactFirstNameIndex, valueFields, contactLastNameIndex);
                            UpdateContactEmailAddress(contactEmailIndex, contact, valueFields);

                            if ((lockAttemptResult.Status != LockAttemptStatus.AlreadyLocked) &&
                                (lockAttemptResult.Status != LockAttemptStatus.NotFound))
                            {
                                if (contact.AutomationStates().IsInEngagementPlan(options.PlanId))
                                {
                                    contact.AutomationStates().MoveToEngagementState(options.PlanId, options.StateId);
                                    Log.Info(
                                        string.Format("Move contact: {0} to engagement plan stateId: {1}",
                                            contact.ContactId, options.StateId), this);
                                }
                                else
                                {
                                    contact.AutomationStates().EnrollInEngagementPlan(options.PlanId, options.StateId);
                                    Log.Info(
                                        string.Format("Enrolled contact: {0} to engagement plan stateId: {1}",
                                            contact.ContactId, options.StateId), this);
                                }

                                contactRepository.SaveContact(contact, new ContactSaveOptions(true, leaseOwner, null));
                                numOfContactsAddedToState++;
                            }
                            else
                            {
                                Log.Error(
                                    string.Format("Failed to enroll contact: {0} in engagement plan stateId: {1}",
                                        contact.ContactId, options.StateId), this);
                            }
                            numOfImported++;
                        }

                        valueFields = csvFileReader.ReadLine();
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex.Message, ex, this);
                }
            }

            Log.Info(
                string.Format(
                    "Import Contacts Finished: Imported: {0}, Contact Exists: {1}, Bad Contact Data: {2}, Added to State: {3}",
                    numOfImported, numOfContactExists, numOfBadContacts, numOfContactsAddedToState), GetType());
            return numOfImported + "|" + numOfContactExists + "|" + numOfBadContacts + "|" + "|" +
                   numOfContactsAddedToState + "|";
        }

        private void UpdateContactPersonalInfo(Contact contact, int contactFirstNameIndex, List<string> valueFields,
            int contactLastNameIndex)
        {
            var contactPersonalInfoFacet = contact.GetFacet<IContactPersonalInfo>("Personal");
            if (contactFirstNameIndex >= 0)
                contactPersonalInfoFacet.FirstName = valueFields[contactFirstNameIndex];

            if (contactLastNameIndex >= 0)
                contactPersonalInfoFacet.Surname = valueFields[contactLastNameIndex];
        }

        private void UpdateContactEmailAddress(int contactEmailIndex, Contact contact, List<string> valueFields)
        {
            if (contactEmailIndex >= 0)
            {
                var contactEmailFacet = contact.GetFacet<IContactEmailAddresses>("Emails");
                contactEmailFacet.Preferred = "Preferred";
                if (contactEmailFacet.Entries.Contains("Preferred"))
                {
                    var preferredEmailAddress = contactEmailFacet.Entries["Preferred"];
                    preferredEmailAddress.SmtpAddress = valueFields[contactEmailIndex];
                }
                else
                {
                    var preferredEmailAddress = contactEmailFacet.Entries.Create("Preferred");
                    preferredEmailAddress.SmtpAddress = valueFields[contactEmailIndex];
                }
            }
        }

        protected class ContactImportOptions
        {
            public string FileName { get; set; }
            public ID PlanId { get; set; }
            public ID StateId { get; set; }
            public string ContactIdentifierField { get; set; }
            public string ContactFirstNameField { get; set; }
            public string ContactLastNameField { get; set; }
            public string ContactEmailField { get; set; }
        }

        protected enum ImportActionType
        {
            Unknown,
            ImportAsUser,
            ImportAsContact
        }

        #region Methods from base class

        protected new bool ReadFile()
        {
            if (string.IsNullOrEmpty(RealFilename))
            {
                SheerResponse.Alert(Translate.Text("Select a file first."));
                return false;
            }
            if (!FileUtil.FileExists(RealFilename))
            {
                SheerResponse.Alert(Translate.Text("'{0}' file does not exist.", (object) RealFilename));
                return false;
            }
            if (!RealFilename.Equals(LastFile))
                using (var csvFileReader = new CsvFileReader(RealFilename))
                {
                    FieldList = csvFileReader.ReadLine();
                    LastFile = RealFilename;
                    InitializeFields();
                }
            return true;
        }

        protected new void CheckImport()
        {
            var str = Context.ClientPage.ServerProperties["job"] as string;
            var job = !string.IsNullOrEmpty(str) ? JobManager.GetJob(Handle.Parse(str)) : null;
            if (job == null)
                Next();
            else if (job.IsDone)
            {
                if (job.Status.Result != null)
                    UpdateForm(job.Status.Result.ToString());
                Next();
            }
            else
                SheerResponse.Timer("CheckImport", 300);
        }

        private void UpdateForm(string results)
        {
            if (string.IsNullOrEmpty(results))
                return;
            var strArray = results.Split('|');
            if (strArray.Length < 5)
                return;
            NumImported.Text = strArray[0];
            NumUserExists.Text = strArray[1];
            NumBadUserName.Text = strArray[2];
            NumBroken.Text = strArray[3];
            NumAddedToState.Text = strArray[4];
            SheerResponse.Refresh(Results);
        }

        private void AddDelSection(string idPostfix)
        {
            Assert.ArgumentNotNull(idPostfix, "idPostfix");
            var imageBuilder = new ImageBuilder
            {
                Src = "Applications/16x16/delete2.png",
                Width = 16,
                Height = 16,
                Style = "cursor: pointer"
            };
            var literal = new Literal
            {
                Text = imageBuilder.ToString()
            };
            var border1 = new Border();
            border1.ID = "DelSection" + idPostfix;
            border1.Click = "DelSection_Click";
            var border2 = border1;
            border2.Controls.Add(literal);
            ParentControl.Controls.Add(border2);
        }

        private void AddFieldCombobox(string idPostfix, string defValue)
        {
            Assert.ArgumentNotNull(idPostfix, "idPostfix");
            var editableCombobox1 = new EditableCombobox();
            editableCombobox1.ID = "Field" + idPostfix;
            editableCombobox1.SelectOnly = true;
            editableCombobox1.List = FieldList;
            var editableCombobox2 = editableCombobox1;
            if (!string.IsNullOrEmpty(defValue))
                foreach (var str in FieldList)
                    if (defValue.Equals(str, StringComparison.OrdinalIgnoreCase))
                    {
                        editableCombobox2.Value = str;
                        break;
                    }
            if (string.IsNullOrEmpty(editableCombobox2.Value))
                editableCombobox2.Value =
                    HttpUtility.HtmlEncode(FormatDefaultText(Translate.Text("select to add field")));
            ParentControl.Controls.Add(editableCombobox2);
        }

        private void AddPropertyCombobox(string idPostfix, string defValue)
        {
            Assert.ArgumentNotNull(idPostfix, "idPostfix");
            var editableCombobox1 = new EditableCombobox();
            editableCombobox1.ID = "Property" + idPostfix;
            editableCombobox1.SelectOnly = true;
            editableCombobox1.List = PropertyList;
            var editableCombobox2 = editableCombobox1;
            if (string.IsNullOrEmpty(defValue))
                editableCombobox2.Value = HttpUtility.HtmlEncode(FormatDefaultText(Translate.Text("select property")));
            else
                foreach (var str in PropertyList)
                    if (defValue.Equals(str, StringComparison.OrdinalIgnoreCase))
                    {
                        editableCombobox2.Value = str;
                        break;
                    }
            ParentControl.Controls.Add(editableCombobox2);
        }

        private string AddRow(string prevIdPostfix)
        {
            return AddRow(prevIdPostfix, string.Empty);
        }

        private string AddRow(string prevIdPostfix, string defProperty)
        {
            var uniqueId = Control.GetUniqueID(string.Empty);
            if (!string.IsNullOrEmpty(prevIdPostfix))
                AddDelSection(prevIdPostfix);
            AddFieldCombobox(uniqueId, defProperty);
            AddPropertyCombobox(uniqueId, defProperty);
            return uniqueId;
        }

        private string FormatDefaultText(string text)
        {
            Assert.ArgumentNotNull(text, "text");
            return "<" + text + ">";
        }

        private void InitializeUserNameField()
        {
            UserName.List = FieldList;
            UserName.Value = FormatDefaultText(Translate.Text("select field"));
        }

        private void StartJob(string name, string method, object helper, params object[] args)
        {
            Assert.ArgumentNotNull(name, "name");
            Assert.ArgumentNotNull(method, "method");
            Assert.ArgumentNotNull(helper, "helper");
            var job = JobManager.Start(new JobOptions(name, method, Context.Site.Name, helper, method, args)
            {
                EnableSecurity = false,
                AfterLife = TimeSpan.FromSeconds(3.0),
                WriteToLog = true
            });
            try
            {
                Context.ClientPage.ServerProperties["job"] = job.Handle.ToString();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex, this);
            }
        }

        #endregion
    }
}