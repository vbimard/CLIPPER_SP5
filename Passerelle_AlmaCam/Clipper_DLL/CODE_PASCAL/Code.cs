    //  Après la cloture : Eclatement du stock et nommage du stock
    // (2018-280-01-01 2018-280-02-01 si deux chutes, etc.)
    //  Les chutes de 2018-280-01-01 seront nommés 2018-280-01-02 2018-280-01-03 etc. etc.
    //  jusqu'à l'utilisation totale de la tôle initiale
    public class MyAfterEndCuttingEvent : AfterEndCuttingEvent
    {
        public override void OnAfterEndCuttingEvent(IContext context, AfterEndCuttingArgs args)
        {
            //  ce code est exécuté autant de fois qu'il y a de multiplicité
            //  fonctionnement identique cloture à la tôle ou au placement
            IEntity nestingToCut = args.ToCutSheetEntity;
            IEntity nesting = context.EntityManager.GetEntity(nestingToCut.GetFieldValueAsInt("_TO_CUT_NESTING"), "_NESTING");
            int nestingMultiplicity = nesting.GetFieldValueAsInt("_QUANTITY");

            IEntityList tolesCoupees = context.EntityManager.GetEntityList("_SHEET", "_SEQUENCED_NESTING", ConditionOperator.Equal, nesting.Id32);
            tolesCoupees.Fill(false);

            foreach (IEntity toleCoupee in tolesCoupees)
            {
                string toleName = toleCoupee.GetFieldValueAsString("_REFERENCE");

                string counterName = toleCoupee.GetFieldValueAsString("_REFERENCE");

                long counter = CounterManager.GetNextCounterValue(context, counterName);

                IEntityList stocks = context.EntityManager.GetEntityList("_STOCK", "_SHEET", ConditionOperator.Equal, toleCoupee.Id32);
                stocks.Fill(false);

                foreach (IEntity stock in stocks)
                {
                    if (!stock.GetFieldValueAsBoolean("AF_STOCK_CFAO"))
                    {
                        if (!stock.GetFieldValueAsBoolean("AF_STOCK_RENAME"))
                        {
                            stock.SetFieldValue("AF_STOCK_NAME", 0.ToString("00"));

                            if (stock.GetFieldValueAsInt("_QUANTITY") == 1)
                            {
                                stock.Delete();
                                IEntityList stockCFAO = context.EntityManager.GetEntityList("_STOCK", LogicOperator.And, "_SHEET", ConditionOperator.Equal, toleCoupee.Id32, "AF_STOCK_CFAO", ConditionOperator.Equal, true);
                                stockCFAO.Fill(false);
                                IEntity toleCFAO = stockCFAO.FirstOrDefault();

                                //Cree une nouvelle ligne de stock
                                toleCFAO.SetFieldValue("AF_STOCK_NAME", counterName + "-" + counter.ToString("00"));
                                toleCFAO.SetFieldValue("AF_STOCK_CFAO", false);
                                toleCFAO.SetFieldValue("AF_STOCK_RENAME", true);
                                toleCFAO.SetFieldValue("_QUANTITY", 1);
                                toleCFAO.Complete = true;
                                toleCFAO.Save();
                            }

                            if (stock.GetFieldValueAsInt("_QUANTITY") > 1)
                            {
                                //  on eclate le stock 1 x qté5 devient 5 x qté1
                                //Décrémente le stock de 1
                                stock.SetFieldValue("_QUANTITY", stock.GetFieldValueAsLong("_QUANTITY") - 1);

                                IEntityList stockCFAO = context.EntityManager.GetEntityList("_STOCK", LogicOperator.And, "_SHEET", ConditionOperator.Equal, toleCoupee.Id32, "AF_STOCK_CFAO", ConditionOperator.Equal, true);
                                stockCFAO.Fill(false);
                                IEntity toleCFAO = stockCFAO.FirstOrDefault();

                                //Cree une nouvelle ligne de stock
                                toleCFAO.SetFieldValue("AF_STOCK_NAME", counterName + "-" + counter.ToString("00"));
                                toleCFAO.SetFieldValue("AF_STOCK_CFAO", false);
                                toleCFAO.SetFieldValue("AF_STOCK_RENAME", true);
                                toleCFAO.SetFieldValue("_QUANTITY", 1);
                                toleCFAO.Complete = true;
                                toleCFAO.Save();
                            }
                            
                            try
                            {
                                stock.Save();
                            }
                            catch
                            {

                            }

                        }
                    }
                }
            }
        }
    }
    
    //  Suppression des stockCFAO à la restauration du placement
    //  Empechait la suppression du placement car le format contenait du stock
    public class MyBeforeNestingRestoreEvent : BeforeNestingRestoreEvent
    {
        public override void OnBeforeNestingRestoreEvent(IContext context, BeforeNestingRestoreArgs args)
        {
            IEntity nesting = args.NestingEntity;

            string nestingReference = nesting.GetFieldValueAsString("_REFERENCE");
            IEntityList _stockList = context.EntityManager.GetEntityList("_STOCK");
            _stockList.Fill(false);
            var stockList = _stockList.Where(p => p.GetFieldValueAsString("AF_STOCK_NAME")!=null && p.GetFieldValueAsString("AF_STOCK_NAME").StartsWith(nestingReference));
            foreach (IEntity stock in stockList)
            {
                stock.Delete();
            }
        }
    }

    //  Création du stock CFAO
    public class MyBeforeSendToWorkshopEvent : BeforeSendToWorkshopEvent
    {
        public override void OnBeforeSendToWorkshopEvent(IContext context, BeforeSendToWorkshopArgs args)
        {
            IEntity nesting = args.NestingEntity;
            string cle = nesting.GetFieldValueAsString("_REFERENCE");

            #region création stock CFAO

            string nestingName = nesting.GetFieldValueAsString("_NAME");
            int nestingMultiplicity = nesting.GetFieldValueAsInt("_QUANTITY");

            int a = 0;
            IEntityList _sheetList = context.EntityManager.GetEntityList("_SHEET");
            _sheetList.Fill(false);
            var sheetList = _sheetList.Where(p => p.GetFieldValueAsString("_NAME").Contains(nestingName));
            foreach (IEntity sheet in sheetList)
            {
                for (int i = 0; i < nestingMultiplicity; i++)
                {
                    IEntity newStock = context.EntityManager.CreateEntity("_STOCK");
                    newStock.SetFieldValue("_SHEET", sheet.Id32);
                    newStock.SetFieldValue("_NAME", i.ToString());
                    newStock.SetFieldValue("AF_STOCK_NAME", sheet.GetFieldValueAsString("_REFERENCE") + "-" + (i + 1).ToString("00"));
                    newStock.SetFieldValue("_QUANTITY", 1);
                    newStock.SetFieldValue("AF_STOCK_CFAO", true);
                    newStock.Save();
                }
            }

            #endregion  création stock CFAO
        }
    }
    
    public class MyOffcutReferenceFormula : OffcutReferenceFormula
    {
        public override string Execute(INestingEngine nestingEngine, INesting nesting, IOffcut offcut)
        {
            string counterKey = string.Format("{0}-{1}", nesting.Reference, offcut.Index.ToString("00"));
            return string.Format("{0}", counterKey);
        }
    }
