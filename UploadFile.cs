using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using OfficeOpenXml;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Crm.Sdk.Messages;

namespace CreateMetadatafromFile.Controllers
{
    public class UploadFileController : Controller
    {
    
    [HttpPost]
        public ActionResult ViewUploadFile(HttpPostedFileBase file, Models.UploadFileModel uploadFileModel) 
        {
            Boolean verboseLogs = false;
            StringBuilder message = new StringBuilder();
            StringBuilder messageError = new StringBuilder();
            StringBuilder table = new StringBuilder();
            Int32 entityCount = 0;
            Int32 attributeCount = 0;
            string fieldSchemaName = string.Empty;
            List<string> entityNameCreatedList = new List<string>();
            List<string> fieldSchemaNameCreatedList = new List<string>();
            Boolean error = false;
            if (ModelState.IsValid)
            {
                try
                {
                    if (Login.IsValid(Login.OrgURL.ToString(), Login.Email.ToString(), sc.Decrypt(Login.Password.ToString())))
                    {
                        //assign the crm service
                        
                        if (file != null && file.ContentLength > 0 && System.IO.Path.GetExtension(file) == ".xlsx") 
                            {
                                try
                                {
                                    //Set <httpRuntime maxRequestLength="x" /> in your web.config, where x is the number of 
                                    //KB allowed for upload. Default is 4KB
                                    var fileName = System.IO.Path.GetFileName(file.FileName);
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        file.InputStream.CopyTo(ms);
                                        byte[] myFile = ms.GetBuffer();
                                        Entity myNote = new Entity(Annotation.EntityLogicalName);
                                        string subject = "Metadata File: " + fileName.ToString();
                                        myNote["subject"] = subject;
                                        Guid myNoteGuid = service.Create(myNote);
                                        if (myNoteGuid != Guid.Empty)
                                        {
                                            using (ExcelPackage package = new ExcelPackage(ms))
                                            {
                                                    foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                                                    {
                                                        DataTable tbl = createDataTablefromExcel(worksheet, true);
                                                        var missing = from c in findMyRequiredFields()
                                                                      where !tbl.Columns.Contains(c)
                                                                      select c;
                                                        missing.ToList();

                                                        if (missing.Count() > 0)
                                                        {
                                                            error = true;
                                                            message.AppendLine(worksheet.Name + " is not valid for import and will be skipped.  It is missing the following required columns: " + String.Join(Environment.NewLine, missing));
                                                        }
                                                        else
                                                        {
                                                            error = false;
                                                            message.AppendLine(String.Format("Data successfully retrieved from excel-worksheet: {0}. Colum-count:{1} Row-count:{2}",
                                                                                    worksheet.Name, tbl.Columns.Count, tbl.Rows.Count));

                                                            var eN = tbl.AsEnumerable().Select(r => r.Field<string>(findMyRequiredFields().First<string>()));
                                                            string entityName = eN.First<string>();
                                                            if (verboseLogs) message.AppendLine("entityname: " + entityName.ToString());

                                                            #region GetAttributes
                                                            int currentRow;
                                                            List<AttributeMetadata> addedAttributes;
                                                            List<CRMLookup> crmLookupList;
                                                            addedAttributes = new List<AttributeMetadata>();
                                                            crmLookupList = new List<CRMLookup>();
                                                            foreach (DataRow row in tbl.Rows)
                                                            {
                                                                string fieldDisplayName = row[findMyRequiredFields()[1]].ToString();
                                                                fieldSchemaName = uploadFileModel.preFix + row[findMyRequiredFields()[2]].ToString();
                                                                if (verboseLogs) message.AppendLine("fieldSchemaName: " + fieldSchemaName);
                                                                string fieldReq = row[findMyRequiredFields()[3]].ToString();
                                                                AttributeRequiredLevel reqLvl;
                                                                if (fieldReq == "Business Required")
                                                                    reqLvl = AttributeRequiredLevel.ApplicationRequired;
                                                                else
                                                                    reqLvl = AttributeRequiredLevel.None;
                                                                string fieldDataType = row[findMyRequiredFields()[4]].ToString();
                                                                string fieldRecordType = row[findMyRequiredFields()[5]].ToString();
                                                                string fieldFormat = row[findMyRequiredFields()[6]].ToString();
                                                                if (fieldFormat == "Text")
                                                                    strFm = StringFormat.Text;
                                                                else if (fieldFormat == "URL")
                                                                    strFm = StringFormat.Url;
                                                                else if (fieldFormat == "Date Only")
                                                                    dtFm = DateTimeFormat.DateOnly;
                                                                string fieldPrecisionLength = row[findMyRequiredFields()[7]].ToString();
                                                                string fieldOther = row[findMyRequiredFields()[8]].ToString();
                                                                if (fieldDataType == "Decimal Number" && fieldOther.Contains("auto"))
                                                                {
                                                                    minLength = Convert.ToDecimal(fieldOtherMin.Substring(0, fieldOtherMin.IndexOf(' ') - 1).Replace(",", ""));
                                                                    if (verboseLogs) message.AppendLine("minLength: " + minLength.ToString());
                                                                    maxLength = Convert.ToDecimal(fieldOtherMax.Substring(fieldOtherMax.IndexOf(' ') + 1, fieldOtherMax.IndexOf(", auto") - (fieldOtherMax.IndexOf(' ') + 1)).Replace(",", ""));
                                                                    if (verboseLogs) message.AppendLine("maxLength: " + maxLength.ToString());
                                                                }
                                                                else 
                                                                {
                                                                    if (fieldDataType == "Single Line of Text")
                                                                        addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, Convert.ToInt32(fieldPrecisionLength), strFm));
                                                                    else if (fieldDataType == "Multiple Lines of Text")
                                                                        addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, Convert.ToInt32(fieldPrecisionLength.Replace(",",""))));
                                                                    else if (fieldDataType == "Date and Time")
                                                                        addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, dtFm));
                                                                    else if (fieldDataType == "Two Options")
                                                                        addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, defaultValue));
                                                                    else if (fieldDataType == "Decimal Number")
                                                                        addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, minLength, maxLength, Convert.ToInt16(fieldPrecisionLength)));
                                                                    else if (fieldDataType == "Lookup")
                                                                        crmLookupList.Add(new CRMLookup { schemaName = fieldSchemaName, displayName = fieldDisplayName, description = "", fieldRecordType = fieldRecordType, lvl = reqLvl, parentEntityDisplayName = entityName, fieldOther = fieldOther });
                                                                    else if (fieldDataType == "Option Set")
                                                                    {
                                                                            var labels = Regex.Split(fieldOther, "\r\n|\r|\n");
                                                                            OptionSetMetadata setupOptionSetMetadata = new OptionSetMetadata
                                                                            {
                                                                                IsGlobal = false,
                                                                                OptionSetType = OptionSetType.Picklist
                                                                            };
                                                                            int optionsValue = 1;
                                                                            foreach (var lbl in labels)
                                                                            {
                                                                                setupOptionSetMetadata.Options.Add(new OptionMetadata(new Label(lbl, 1033), optionsValue));
                                                                                optionsValue++;
                                                                            }
                                                                            addedAttributes.Add(createAttributeObj(fieldSchemaName, fieldDisplayName, "", reqLvl, setupOptionSetMetadata));
                                                                    }
                                                                }
                                                                currentRow++;
                                                            }
                                                            #endregion

                                                            #region CreateEntityandPrimaryAttribute
                                                            if (verboseLogs) message.AppendLine(String.Format("Create entity: {0} with PrimaryAttribute: {1}", entityName, primaryName));
                                                            CreateEntityRequest createrequest = new CreateEntityRequest
                                                            {
                                                                Entity = new EntityMetadata
                                                                {
                                                                    SchemaName = uploadFileModel.preFix + entityName,
                                                                    DisplayName = new Label(entityName, 1033),
                                                                    DisplayCollectionName = new Label(entityName, 1033)
                                                                    OwnershipType = OwnershipTypes.UserOwned,
                                                                },
                                                                PrimaryAttribute = new StringAttributeMetadata
                                                                {
                                                                    SchemaName = primaryName,
                                                                    RequiredLevel = new AttributeRequiredLevelManagedProperty(primaryLvl),
                                                                    MaxLength = primaryLength,
                                                                    Format = primaryFormat,
                                                                    DisplayName = new Label(primaryDisplayName, 1033)
                                                                },
                                                            };
                                                                service.Execute(createrequest);
                                                                entityCount++;
                                                                attributeCount++;
                                                                if (verboseLogs) message.AppendLine("The " + entityName + " custom entity has been created.");
                                                                entityNameCreatedList.Add(entityName);
                                                            #endregion


                                                            #region CreateAttributes
                                                            if (addedAttributes.Count > 0)
                                                            {
                                                                foreach (AttributeMetadata anAttribute in addedAttributes)
                                                                {
                                                                    if (verboseLogs) message.AppendLine(String.Format("attribute {0}.", anAttribute.SchemaName));
                                                                    fieldSchemaName = anAttribute.SchemaName;
                                                                    CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                                                    {
                                                                        EntityName = uploadFileModel.preFix + entityName.ToLower(),
                                                                        Attribute = anAttribute
                                                                    };
                                                                        service.Execute(createAttributeRequest);
                                                                        attributeCount++;
                                                                        fieldSchemaNameCreatedList.Add(anAttribute.SchemaName);
                                                                }
                                                            }

                                                            if (crmLookupList.Count > 0)
                                                            {
                                                                foreach (CRMLookup anAttribute in crmLookupList)
                                                                {
                                                                    if (verboseLogs) message.AppendLine(String.Format("LookupAttribute {0}.", anAttribute.schemaName));
                                                                    fieldSchemaName = anAttribute.schemaName;
                                                                    CreateOneToManyRequest req = new CreateOneToManyRequest()
                                                                    {
                                                                        Lookup = new LookupAttributeMetadata()
                                                                        {
                                                                            DisplayName = new Label(anAttribute.displayName, 1033),
                                                                            LogicalName = anAttribute.schemaName,
                                                                            SchemaName = anAttribute.schemaName,
                                                                            RequiredLevel = new AttributeRequiredLevelManagedProperty(anAttribute.lvl)
                                                                        },
                                                                        OneToManyRelationship = new OneToManyRelationshipMetadata()
                                                                        {
                                                                            ReferencedEntity = anAttribute.fieldRecordType.ToLower(),
                                                                            ReferencedAttribute = anAttribute.fieldRecordType.ToLower() + "id",
                                                                            ReferencingEntity = uploadFileModel.preFix + anAttribute.parentEntityDisplayName.ToLower(),
                                                                            SchemaName = anAttribute.fieldOther
                                                                        }
                                                                    };
                                                                        service.Execute(req);
                                                                        attributeCount++;
                                                                        fieldSchemaNameCreatedList.Add(anAttribute.schemaName);
                                                                        if (verboseLogs) message.AppendLine(String.Format("Created the LookupAttribute {0}.", anAttribute.schemaName));
                                                                }
                                                            }
                                                            #endregion

                                                            Entity myNoteU = new Entity(Annotation.EntityLogicalName);
                                                            myNoteU.Id = myNoteGuid;
                                                            if (error)
                                                                myNoteU["subject"] = "FAILED " + subject;
                                                            else
                                                                myNoteU["subject"] = "SUCCEEDED " + subject;
                                                            myNoteU["notetext"] = message.ToString() +
                                                                String.Format("Entities created: {0}, Attributes created: {1}, Global OptionSets created: {2}", entityCount, attributeCount, globalAttributeCount);
                                                            service.Update(myNoteU);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    error = true;
                                    messageError.AppendLine(String.Format("Exception: {0}, fieldSchemaName: {1}.  Please try again after adjusting the excel worksheet.", ex.Message, fieldSchemaName));
                                }
                            }
                            else
                            {
                                error = true;
                                messageError.AppendLine("You must select a non-empty xlsx file");
                            }
                        }
                        else
                    {
                        error = true;
                        messageError.AppendLine("Your credentials have expired. You must sign in again");
                    }

                    message.AppendLine(String.Format("Entity count: {0}, Attribute count: {1}", entityCount, attributeCount));
                    message.AppendLine(String.Format("Entities created: {0}, Attributes created: {1}", String.Join(Environment.NewLine, entityNameCreatedList), String.Join(Environment.NewLine, fieldSchemaNameCreatedList)));
                    ViewBag.Error = messageError;
                    ViewBag.Message += message; 
                        
                    if (!error)    
                        ViewBag.Message += "Please review the entities and attributes created within your organization and add to the forms 
                        and views as needed.";
                }
                catch (Exception ex)
                {
                    error = true;
                    messageError.AppendLine(String.Format("Exception: {0}.  Please log in and try again.", ex.Message));
                }
            }
            return View();
        }
        
        private static List<string> findMyRequiredFields()
        {
            List<string> myRequiredFields = new List<string>();
            myRequiredFields.Add("Entity");
            myRequiredFields.Add("FieldName");
            myRequiredFields.Add("FieldSchema");
            myRequiredFields.Add("FieldRequirement");
            myRequiredFields.Add("DataType");
            myRequiredFields.Add("RecordType");
            myRequiredFields.Add("Format");
            myRequiredFields.Add("PrecisionOrLength");
            myRequiredFields.Add("Other");
            return myRequiredFields;
        }
    
    }
}
