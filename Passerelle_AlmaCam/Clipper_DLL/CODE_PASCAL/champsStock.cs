using System;
using System.IO;
using System.Collections.Generic;
using Wpm.Schema.Kernel;
using Wpm.Implement.Manager;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

namespace Wpm.Implement.ModelSetting
{
    public partial class ImportUserEntityType : ScriptModelCustomization, IScriptModelCustomization
    {
        public override bool Execute(IContext context, IContext hostContext)
        {
            #region Stock
            
            {
                IEntityType entityType = context.Kernel.GetEntityType("_STOCK");
                IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(context, 1, entityType , null, "_SUPPLY", null);
                entityTypeFactory.Key = "_STOCK";
                entityTypeFactory.Name = "Stock";
                entityTypeFactory.DefaultDisplayKey = "_NAME";
                entityTypeFactory.ActAsEnvironment = false;
                
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_RENAME";
                    fieldDescription.Name = "Renommé ?";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_CFAO";
                    fieldDescription.Name = "Stock CFAO ?";
                    fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                    fieldDescription.DefaultValue = false;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                {
                    IFieldDescription fieldDescription = new FieldDescription(context.Kernel.UnitSystem, true);
                    fieldDescription.Key = "AF_STOCK_NAME";
                    fieldDescription.Name = "Nom stock";
                    fieldDescription.Editable = FieldDescriptionEditableType.NoEditable;
                    fieldDescription.Visible = FieldDescriptionVisibilityType.AllSection;
                    fieldDescription.Mandatory = false;
                    fieldDescription.FieldDescriptionType = FieldDescriptionType.String;
                    entityTypeFactory.EntityTypeAttributList.Add(fieldDescription);
                    
                }
                if (!entityTypeFactory.UpdateModel())
                {
                    foreach (ModelSettingError error in entityTypeFactory.ErrorList)
                    {
                        hostContext.TraceLogger.TraceError(error.Message, true);
                    }
                    return false;
                }
                
            }
            
            #endregion
            return true;
        }
    }
}
