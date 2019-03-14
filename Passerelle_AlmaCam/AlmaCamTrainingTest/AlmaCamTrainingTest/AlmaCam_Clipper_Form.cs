using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Wpm.Implement.Manager;
using Wpm.Implement.Processor;
using Wpm.Implement.ComponentEditor;  // ouverture de fenetres de selection almacam
using Wpm.Schema.Kernel;

//using System.ComponentModel;
//using Alma.BaseUI.Utils;





//dll personnalisées
using AF_Clipper_Dll;
using AF_ImportTools;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using Wpm.Implement.ModelSetting;
using Alma.BaseUI.DescriptionEditor;

//suppression de wpm.schema.component.dll et remplacement par wpm.schema.componenteditor.dll 

namespace AlmaCamTrainingTest
{

   

    public partial class AlmaCam_Clipper_Form : Form,IDisposable
    {

        //initialisation des listes
        IContext _Context = null;
        
        string DbName = Alma_RegitryInfos.GetLastDataBase();

        public AlmaCam_Clipper_Form()
        {
            try
            {

                InitializeComponent();
               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }
        /// <summary>
        /// icone notification,
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

       
       

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            AF_ImportTools.SimplifiedMethods.NotifyMessage("ALMACAM CLIPPER", "Starting Almacam Clipper",20000);
            //creation du model repository
            var databases = new List<string>();
            IModelsRepository modelsRepository = new ModelsRepository();
             foreach (IModel model in modelsRepository.ModelList)
            {
                databases.Add(model.DatabaseName);

            }
            //nom de la base;
            Lst_Model.DataSource = databases;
            Lst_Model.SelectedItem = DbName;
            string infosPasserelle;

            //

            //_Context = modelsRepository.GetModelContext(DbName);
            //int i = _Context.ModelsRepository.ModelList.Count();
            SimplifiedMethods.NotifyMessage("AlmaCam", " Selected " + DbName);
            infosPasserelle = DbName + "-P." + AF_Clipper_Dll.Clipper_Param.GetClipperDllVersion() + "-CAM." + AF_Clipper_Dll.Clipper_Param.GetAlmaCAMCompatibleVerion();
            this.Text = this.Name;
            this.InfosLabel.Text = infosPasserelle;
            this.Text = "Passerelle Clipper V8 validée pour : " + infosPasserelle;
        }

        #region button
        private void button1_Click(object sender, EventArgs e)
        {//purge stock
            IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
            stocks.Fill(false);
            DialogResult res = MessageBox.Show("Do you really want to destroy all sheets from the stock?", "Warnig", MessageBoxButtons.OKCancel);
            if (res == DialogResult.OK)
            {
                foreach (IEntity stock in stocks)
                {
                    stock.Delete();
                }


                IEntityList formats = _Context.EntityManager.GetEntityList("_SHEET");
                formats.Fill(false);

                foreach (IEntity format in formats)
                {
                    format.Delete();
                }


                MessageBox.Show("Terminé");

            }
            _Context = null;
        }
    
        private void button3_Click(object sender, EventArgs e)
        {
            
            using (Clipper_Stock Stock = new Clipper_Stock())
            {
                //Stock.Import(_Context, csvImportPath, dataModelstring);
                Stock.Import(_Context);//, csvImportPath);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {

            //import of
            //chargement de sparamètres
            // bool SansDt=false;

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";
          

            string dataModelstring = Clipper_Param.GetModelCA();


            using (Clipper_OF CahierAffaire = new Clipper_OF())
            {
                CahierAffaire.Import(_Context, csvImportPath, dataModelstring, false);
                CahierAffaire.Import(_Context, csvImportSandDt, dataModelstring, true);
                //}

            }


        }
        
        private void button5_Click(object sender, EventArgs e)
        {

            /*
            //private void button8_Click(object sender, EventArgs e)

            Clipper_Dll.Clipper_DoOnAction_AfterSendToWorkshop doonaction = new Clipper_DoOnAction_AfterSendToWorkshop();
            doonaction.execute(_Context);
            //doonaction.OnAfterSendToWorkshopEvent(_Context ,"" );
            */


        }

        private void button6_Click(object sender, EventArgs e)
        {
            ///for test
            ///
            ///
            ///Pour la passerelle « ALMACAM », je vous fournirai un executable pour l’import automatique du stock et du cahier d’affaire en suivant la meme logique.
            ///Chemin de l’executable paramétable et paramétré par defaut dans
            ///

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_DM");
            ProcessStartInfo start_dm = new ProcessStartInfo();
            start_dm.Arguments = "stock " + csvImportPath;
            //start.FileName =  @"C:\AlmaCAM\Bin\AlmaCamUser1.exe";
            start_dm.FileName = Clipper_Param.Get_application1();
            //start.WindowStyle = ProcessWindowStyle.Normal;
            start_dm.CreateNoWindow = true;
            start_dm.UseShellExecute = true;
            System.Diagnostics.Process.Start(start_dm);


        }

        private void button7_Click(object sender, EventArgs e)
        {
            Clipper_Param.GetlistParam(_Context);

            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            ProcessStartInfo start = new ProcessStartInfo();
            ProcessStartInfo start_ca = new ProcessStartInfo();
            start_ca.Arguments = "OF " + csvImportPath;
            //start.FileName = @"C:\AlmaCAM\Bin\Clipper_Import.exe";
            start_ca.FileName = Clipper_Param.Get_application1();
            //start.WindowStyle = ProcessWindowStyle.Normal;
            start_ca.CreateNoWindow = false;
            start_ca.UseShellExecute = true;
            string exename = Clipper_Param.Get_application1();
            Process p = Process.Start(exename, "OF " + csvImportPath);


        }

        private void button8_Click(object sender, EventArgs e)
        {
            /*
            Clipper_Dll.Clipper_DoOnAction_BeforeSendToWorkshop doonaction = new Clipper_DoOnAction_BeforeSendToWorkshop();
            doonaction.OnBeforeSendToWorkshopEvent(_Context ,null );
            */


        }

        private void button9_Click(object sender, EventArgs e)
        {
            //initialisation des listes
            //IContext _Context;
            //Param_Clipper.context = Context;
           

            AF_Clipper_Dll.Clipper_Export_DT Export_dt = new AF_Clipper_Dll.Clipper_Export_DT();
            Export_dt.execute(_Context);
            /*
            Clipper_Dll.Clipper_RemonteeDt Remontee_Dt = new Clipper_Dll.Clipper_RemonteeDt();
            Remontee_Dt.Export_Piece_To_File(_Context);*/
        }

        private void button10_Click(object sender, EventArgs e)
        {///creation devis
            //Clipper_Dll.Test_quote testquote = new Clipper_Dll.Test_quote();
            //testquote.createnewquote(_Context);
            IEntity mpart = null;
            IEntityList mparts = _Context.EntityManager.GetEntityList("_MACHINABLE_PART");
            mparts.Fill(false);
            mpart = mparts.FirstOrDefault();
            Topologie topo = new Topologie();
            topo.GetTopologie(ref mpart, ref _Context);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            long decalage;
            decalage = _Context.ParameterSetManager.GetParameterValue("_EXPORT", "_CLIPPER_QUOTE_NUMBER_OFFSET").GetValueAsLong();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            //appel de la lib d'export des besoins ici
            /*
            using (Clipper_Sheet_Requirement Sheet_requirement = new Clipper_Sheet_Requirement())
            {
                Sheet_requirement.Export(_Context);//), csvImportPath);
            }*/

        }
        #endregion
        #region menu
     
        private void donnéesTechniquesToolStripMenuItem_Click(object sender, EventArgs e)
        {

            AF_Clipper_Dll.Clipper_Export_DT Export_dt = new AF_Clipper_Dll.Clipper_Export_DT();
            Export_dt.Execute();

        }

        private void stockToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            //creation du model repository
       

            using (Clipper_Stock Stock = new Clipper_Stock())
            {
                IModelsRepository modelsRepository = new ModelsRepository();
                IContext myContext = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

                //Stock.Import(_Context, csvImportPath, dataModelstring);
                Stock.Import(_Context);//, csvImportPath);
            }




        }
        private IEntity SelectEntity(string SelectionType)
        {
            IEntity selectedentity = null;
            IEntitySelector xselector = null;
            xselector = new EntitySelector();

            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            xselector.Init(_Context, _Context.Kernel.GetEntityType(SelectionType));//"_TO_CUT_NESTING"
            xselector.MultiSelect = true;

            if (xselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity xentity in xselector.SelectedEntity)
                {
                    selectedentity = xentity;

                }

            }

            return selectedentity;

        }
        private IEntity SelectEntity_If_Not_Exported(string entitystage)
        {
            IEntity selectedentity = null;
            IEntitySelector xselector = null;
            xselector = new EntitySelector();

            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            //GPAO_Exported
            //_Context.EntityManager.GetEntityList("_CLOSED_NESTING"entitystage, "GPAO_Exported", ConditionOperator.Equal, null);
            IEntityList GPAO_Exported_filter = _Context.EntityManager.GetEntityList(entitystage, "GPAO_Exported", ConditionOperator.Equal, null);
            GPAO_Exported_filter.Fill(false);

            _Context.EntityManager.GetExtendedEntityList(entitystage, GPAO_Exported_filter);
            IDynamicExtendedEntityList exported_nestings = _Context.EntityManager.GetDynamicExtendedEntityList(entitystage, GPAO_Exported_filter);
            exported_nestings.Fill(false);

            if (exported_nestings.Count == 0)
            {
                //_Context.EntityManager.GetEntityList("_CLOSED_NESTING"entitystage, "GPAO_Exported", ConditionOperator.Equal, false);
                GPAO_Exported_filter = _Context.EntityManager.GetEntityList(entitystage, "GPAO_Exported", ConditionOperator.Equal, false);
                GPAO_Exported_filter.Fill(false);
               // _Context.EntityManager.GetExtendedEntityList("_CLOSED_NESTING"entitystage, GPAO_Exported_filter);
                exported_nestings = _Context.EntityManager.GetDynamicExtendedEntityList(entitystage, GPAO_Exported_filter);
                exported_nestings.Fill(false);
            }

            xselector.Init(_Context, exported_nestings);//"_TO_CUT_NESTING"
            xselector.MultiSelect = true;

            if (xselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity xentity in xselector.SelectedEntity)
                {
                    selectedentity = xentity;

                }

            }

            return selectedentity;
        
    
              /*
            else
            {
                xselector.Init(_Context, _Context.Kernel.GetEntityType(entitytype));//"_TO_CUT_NESTING"
                xselector.MultiSelect = true;

            }*/

          

            

        }

        private void quitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
        }

        private void quitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void purgerStockToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try {

                IModelsRepository modelsRepository = new ModelsRepository();
                _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
                stocks.Fill(false);
                DialogResult res = MessageBox.Show("Do you really want to destroy all sheets from the stock?", "Warnig", MessageBoxButtons.OKCancel);
                if (res == DialogResult.OK)
                {
                    foreach (IEntity stock in stocks)
                    {
                        stock.Delete();
                    }


                    IEntityList formats = _Context.EntityManager.GetEntityList("_SHEET");
                    formats.Fill(false);

                    foreach (IEntity format in formats)
                    {
                        format.Delete();
                    }

                }



            } catch (Exception ie)
            {
                MessageBox.Show(ie.Message);


            }
            finally { _Context = null; }

          
        }

        private void importStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void relanceClotureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void suppressionStockClotureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string stage = "_CLOSED_NESTING";
           
            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    StockManager.DeleteAlmaCamStock(nesting);

                }
            }


              _Context = null; 
        }

        private void purgerToutLeStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                IModelsRepository modelsRepository = new ModelsRepository();
                _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                //purge stock
                IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
                stocks.Fill(false);
                DialogResult res = MessageBox.Show("Do you really want to destroy all unused sheets from the stock?", "Warning", MessageBoxButtons.OKCancel);
                if (res == DialogResult.OK)
                {
                    foreach (IEntity stock in stocks)
                    {
                        if (stock.GetFieldValueAsEntity("_SHEET").GetFieldValueAsLong("_IN_PRODUCTION_QUANTITY") == 0)
                        {
                            //stock.Delete();

                            if (stock.GetFieldValueAsLong("_USED_QUANTITY") == 0)
                            {
                                stock.Delete();
                            }
                            else if (stock.GetFieldValueAsLong("_BOOKED_QUANTITY") == 0)
                            {
                                stock.Delete();
                            }

                        }
                    }

                    //suppression formats
                    IEntityList formats = _Context.EntityManager.GetEntityList("_SHEET");
                    formats.Fill(false);

                    foreach (IEntity format in formats)
                    {
                        if (format.GetFieldValueAsLong("_IN_PRODUCTION_QUANTITY") == 0)
                        {  
                            format.Delete();
                        }

                    }

                }

                MessageBox.Show("Fin...");
            }



            catch (Exception ie)
            {
                MessageBox.Show(ie.Message);

            }
            finally { _Context = null; }
        }

        private void importerStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            using (var Stock = new Clipper_8_Import_Stock())
            {
                Stock.Import(_Context);//, csvImportPath);
            }

                  _Context = null;

        }

        private void importOFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //import of
            //chargement de sparamètres
            // bool SansDt=false;

        }

        private void avecDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";


            string dataModelstring = Clipper_Param.GetModelCA();


            using (var CahierAffaire = new Clipper_8_Import_OF())
            {

                CahierAffaire.Import(_Context, csvImportPath, dataModelstring, false);
            
            }
             _Context = null; 
        }

        private void sansDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";
            
            string dataModelstring = Clipper_Param.GetModelCA();


            using (var CahierAffaire = new Clipper_8_Import_OF())
            {

                CahierAffaire.Import(_Context, csvImportSandDt, dataModelstring, true);
                //}

            }
              _Context = null; 
        }

        private void relanceClotureToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;

        }

        private void relanceenvoiecoupeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            var doonaction = new Clipper_8_DoOnAction_AfterSendToWorkshop();
            string stage = "_TO_CUT_NESTING";

            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    doonaction.Execute(nesting);

                }
            }
              _Context = null; 
        }

        private void reinitialiserStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IEntity TO_CUT_nesting;
           
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            var doonaction = new Clipper_8_Before_Nesting_Restore_Event();
            
            string stage = "_TO_CUT_NESTING";

            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector nestingselector = null;

            nestingselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            nestingselector.MultiSelect = true;


            if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (IEntity nesting in nestingselector.SelectedEntity)
                {
                    doonaction.Execute(nesting);

                }
            }
             _Context = null;         }

        private void recreerLEmfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            List<IEntity> stocks = SimplifiedMethods.Get_Entity_Selector(_Context, "_STOCK");
            if (stocks.Count>0)
            {
                MessageBox.Show("Seules les Chutes pourront obtenir de nouveaux emf");
            

                        foreach (IEntity currentstockentity in stocks)
                        {
                            // 
                            if (currentstockentity.GetFieldValueAsEntity("_SHEET").GetFieldValueAsInt("_TYPE") == 2) {
                                string fieldvalue = currentstockentity.GetFieldValueAsString("FILENAME");
                                //on ne remplace que les chaine filname vides.
                                if (String.IsNullOrEmpty(fieldvalue)==true){ 
                                    string emffilepath = StockManager.Create_EMF_Of_Stock_Entity(currentstockentity);
                                    currentstockentity.SetFieldValue("FILENAME", emffilepath);
                                    currentstockentity.Save();
                                }
                            }
                //
            }

            }
            _Context = null; 




        }

        private void relanceClotureToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);


            SimplifiedMethods.NotifyMessage("Almacam_Clipper", "Ouverture de la base de " + Lst_Model.Text + " entités de stock.");

                   
            Cursor.Current = Cursors.WaitCursor;





            List<IEntity> cutsheets = SimplifiedMethods.Get_Entity_Selector(_Context, "_CUT_SHEET");
            var doonaction = new Clipper_8_DoOnAction_After_Cutting_end();
            foreach (IEntity cutsheet in cutsheets)
            {
                // 
                doonaction.Execute(cutsheet);
                //
            }

            Cursor.Current = Cursors.Default;
            _Context = null; 

        }

        private void preparerLaBasePourClipperToolStripMenuItem_Click(object sender, EventArgs e)
        {




            //IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

            IModelsRepository modelsRepository = new ModelsRepository();
         
            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();
          
            CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.Entities3));
            CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources._event));
            CustoFileCsList.Add(new Tuple<string, string>("Commandes", Properties.Resources.CommandesV2));

            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);
                
                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                 File.Delete(CommandCsPath);
            }
           

        }

        private void iMportChaierDaffaireToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            Clipper_Param.GetlistParam(_Context);
            string csvImportPath = Clipper_Param.GetPath("IMPORT_CDA");
            //recuperation du nom de fichier
            string csvFileName = Path.GetFileNameWithoutExtension(csvImportPath);
            string csvDirectory = Path.GetDirectoryName(csvImportPath);
            string csvImportSandDt = csvDirectory + "\\" + csvFileName + "_SANSDT.csv";


            string dataModelstring = Clipper_Param.GetModelCA();


            using (Clipper_OF CahierAffaire = new Clipper_OF())
            {
                CahierAffaire.Import(_Context, csvImportPath, dataModelstring, false);
                CahierAffaire.Import(_Context, csvImportSandDt, dataModelstring, true);
                //}

            }
             _Context = null; 
        }

        private void c7ImpoortStockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //creation du model repository
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);

            using (Clipper_Stock Stock = new Clipper_Stock())
            {
                //Stock.Import(_Context, csvImportPath, dataModelstring);
                Stock.Import(_Context);//, csvImportPath);
            }

            MessageBox.Show("import Stock terminé");
            _Context = null; 
        }

        private void remonterDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            AF_Clipper_Dll.Clipper_Export_DT Export_dt = new AF_Clipper_Dll.Clipper_Export_DT();
            Export_dt.execute(_Context);
            _Context = null; 

        }

        private void creerRetourGPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IModelsRepository modelsRepository = new ModelsRepository();
            _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            AF_Clipper_Dll.Clipper_DoOnAction_After_Cutting_end doonaction = new Clipper_DoOnAction_After_Cutting_end();


            string stage = "_CUT_SHEET";
            //creation du fichier de sortie
            //recupere les path
            Clipper_Param.GetlistParam(_Context);
            IEntitySelector Entityselector = null;

            Entityselector = new EntitySelector();

            //entity type pointe sur la list d'objet du model
            Entityselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
            Entityselector.MultiSelect = true;

            if (Entityselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                doonaction.execute(Entityselector.SelectedEntity);
            }

              _Context = null; 
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

        private void extraireChampsClipperToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //selection database //
            //LogMessage.Write("Launch all Install Customisation");


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void ajusterLesTolesIsGpaoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ///mise a jour des champs nnon existant
          
        }

       
        private void initialiserStockToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                SimplifiedMethods.NotifyMessage("Almacam_Clipper", "Ouverture de la base de " + Lst_Model.Text + " entités de stock.");

                //creation du model repository
                IModelsRepository modelsRepository = new ModelsRepository();
                IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;
                                                                                        //set value 
                IEntityList stocks = _Context.EntityManager.GetEntityList("_STOCK");
                stocks.Fill(false);


                SimplifiedMethods.NotifyMessage("Almacam_Clipper", "traitement de " + stocks.Count() + " entités de stock.");

                Cursor.Current = Cursors.WaitCursor;

                long position = 0;
                int step = 1;
                foreach (IEntity stock in stocks)
                {
                    position++;

                    SimplifiedMethods.NotifyStatusMessage("", " Initialisation   en cours ...", stocks.Count(), position,ref step, 0);

                    /*
                    if (stock.GetFieldValueAsLong("_QUANTITY") == 0 && stock.GetFieldValueAsLong("_BOOKED_QUANTITY") == 0)
                    { stock.SetFieldValue("AF_IS_OMMITED", true); }
                    else
                    { stock.SetFieldValue("AF_IS_OMMITED", false); }
                    */
                    //les données seront settées par clipper


                    stock.SetFieldValue("AF_IS_OMMITED", false);
                    if (stock.GetFieldValueAsString("IDCLIP") == string.Empty)

                    {// non ttraite
                        stock.SetFieldValue("IDCLIP", null);
                        //stock.SetFieldValue("AF_IS_OMMITED", true);

                    }
                    //recuperation du champs filename de la tole dans le champs filename du stock

                    ///recupération du filename des sheet dans le stock
                    string filname = stock.GetFieldValueAsEntity("_SHEET").GetFieldValueAsString("FILENAME");

                    if (filname != string.Empty || filname != null)
                    {
                        stock.SetFieldValue("FILENAME", filname);
                        //Creating_emf;
                        if (File.Exists(filname) == false)
                        {
                            //create emf //

                        }
                    }




                    stock.Save();
                }
                SimplifiedMethods.NotifyMessage("Almacam_Clipper", "Operation terminée ");
                // Execute your time-intensive hashing code here...

                // Set cursor as default arrow
                
                //MessageBox.Show("Operation terminée");
            }

            finally
            {
                  _Context = null; 
                Cursor.Current = Cursors.Default;
            }
        
        }





        #endregion

        private void remonteeDTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //creation du model repository
            IModelsRepository modelsRepository = new ModelsRepository();
            IContext _Context = modelsRepository.GetModelContext(Lst_Model.Text);  //nom de la base;

            var Export_dt = new AF_Clipper_Dll.Clipper_8_Export_DT();
            Export_dt.Execute(_Context);

              _Context = null; 
        }

        private void c7PlanningToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
                //IEntity TO_CUT_nesting;
                IModelsRepository modelsRepository = new ModelsRepository();
                _Context = modelsRepository.GetModelContext(Lst_Model.Text);

                var doonaction = new Clipper_DoOnAction_AfterSendToWorkshop_ForPlanning();
                string stage = "_TO_CUT_NESTING";
           
                //creation du fichier de sortie
                //recupere les path
                Clipper_Param.GetlistParam(_Context);
                IEntitySelector nestingselector = null;

                nestingselector = new EntitySelector();

                //entity type pointe sur la list d'objet du model
                nestingselector.Init(_Context, _Context.Kernel.GetEntityType(stage));
                nestingselector.MultiSelect = true;


                if (nestingselector.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (IEntity nesting in nestingselector.SelectedEntity)
                    {
                        doonaction.Execute(nesting);

                    }
                }
                _Context = null;
            
        }

        private void ajouterLesEvenementsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();

         
            CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources._event));
          

            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }

        private void ajouterLesCommandesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();

            //CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.Entities3));
            //CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources._event));
            CustoFileCsList.Add(new Tuple<string, string>("Commandes", Properties.Resources.CommandesV2));

            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }

        private void ajouerLesChampsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //IList<CustomizedCommandItem> customizedItemList = GetCustomizedCommandTypeList(context);

            IModelsRepository modelsRepository = new ModelsRepository();

            List<Tuple<string, string>> CustoFileCsList = new List<Tuple<string, string>>();

            CustoFileCsList.Add(new Tuple<string, string>("Entities", Properties.Resources.Entities3));
            //CustoFileCsList.Add(new Tuple<string, string>("FormulasAndEvents", Properties.Resources._event));
            //CustoFileCsList.Add(new Tuple<string, string>("Commandes", Properties.Resources.CommandesV2));

            foreach (Tuple<string, string> CustoFileCs in CustoFileCsList)
            {
                string CommandCsPath = Path.GetTempPath() + CustoFileCs.Item1 + ".cs";
                File.WriteAllText(CommandCsPath, CustoFileCs.Item2);

                string cs = CustoFileCs.Item1;

                ModelManager modelManager = new ModelManager(modelsRepository);
                modelManager.CustomizeModel(CommandCsPath, Lst_Model.Text, true);
                File.Delete(CommandCsPath);
            }
        }

        private void ajouterLesChampsSpécifiquesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }

    public class NewField
    {
        public string Key;
        public string Name;
        public FieldDescriptionEditableType Editable;
        public FieldDescriptionVisibilityType Visible;
        public FieldDescriptionType Type;
        public string Link;
        public NewField(string key, string name, FieldDescriptionEditableType editable, FieldDescriptionVisibilityType visible, FieldDescriptionType type, string link)
        {
            Key = key;
            Name = name;
            Editable = editable;
            Visible = visible;
            Type = type;
            Link = link;
        }
    }


    public class InstallclipperTest : ScriptModelCustomization, IScriptModelCustomization
    {
        IContext _Context;
        
        IContext _HostContext ;

        public InstallclipperTest(IContext context){
          
            IModelsRepository modelsRepository = new ModelsRepository();
            _HostContext = modelsRepository.GetModelContext("ModelsRepository");  //nom de la base;
            //_HostContext = modelsRepository.GetModelContext("ModelsRepositoryAlmaCAM");  //nom de la base;
            _Context = context;

        }
        
            public override bool Execute(IContext _Context, IContext hostContext)
            {


                #region Stock

                {
                    IEntityType entityType = _Context.Kernel.GetEntityType("_STOCK");
                    IEntityTypeFactory entityTypeFactory = new EntityTypeFactory(_Context, 1, entityType, null, "_SUPPLY", null);
                    entityTypeFactory.Key = "_STOCK";
                    entityTypeFactory.Name = "Stock";
                    entityTypeFactory.DefaultDisplayKey = "_NAME";
                    entityTypeFactory.ActAsEnvironment = false;

                    //if(entityType.FieldList)

                    {
                        IFieldDescription fieldDescription = new FieldDescription(_Context.Kernel.UnitSystem, true);
                        fieldDescription.Key = "FILENAME";
                        fieldDescription.Name = "*Renommé ?";
                        fieldDescription.Editable = FieldDescriptionEditableType.AllSection;
                        fieldDescription.Visible = FieldDescriptionVisibilityType.Invisible;
                        fieldDescription.Mandatory = false;
                        fieldDescription.FieldDescriptionType = FieldDescriptionType.Boolean;
                        fieldDescription.DefaultValue = false;
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











