﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.Odbc;
using System.IO;
using Wpm.Implement.Processor;
using AF_ImportTools;
using Wpm.Implement.Manager;
using Wpm.Schema.Kernel;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Win32;
using System.Runtime.Serialization;
using System.Data.Common;
using Actcut.CommonModel;

namespace AF_Import_ODBC_Clipper_AlmaCam

{
   
    #region  commandes
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// creation des commandes d'imports
    /// </summary>
    //section des commandes
    /// <summary>
    /// bouton d'import des matieres
    /// </summary>
    public class Clipper8_ImportMatiere_Processor : CommandProcessor
    {

        public override bool Execute()
        {
            
            try {

                
                Import(Context);
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }
           
        }

        /////////////////////////////////////////////////////////////////
        /// <summary>
        /// import des entité clipper
        /// </summary>
        /// <param name="contextcontextlocal"></param>
        /// 
        public void Import(IContext contextcontextlocal)
        {
           
            TextWriterTraceListener logFile;
            try
            {
                //creation des logs
                using (Clipper_Import_Matiere Update_Material = new Clipper_Import_Matiere())
                {
                    try
                    {

                        Cursor.Current = Cursors.WaitCursor;

                        Update_Material.Import(Context);

                        Cursor.Current = Cursors.Default;


                    }
                    catch { Cursor.Current = Cursors.Default; }
                }


            }
            catch(Exception ie)
            {
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
               
            }
        }
    }

    /// <summary>
    /// bouton d'import des toles
    /// attention l'import des toles est fait par l'import du stock clipper pour le moment.
    /// </summary>
    public class Clipper_Import_Toles_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                //creation des logs
                TextWriterTraceListener logFile;
                Import(Context);
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }

        }
        public bool Import(IContext contextlocal)
        {
            try
            {
                using (Clipper_Article_Toles_MonoDim Toles_MonoDim = new Clipper_Article_Toles_MonoDim(contextlocal))
                {
                    try
                    {


                        Cursor.Current = Cursors.WaitCursor;


                        Toles_MonoDim.Read();
                        Toles_MonoDim.Write();
                        Toles_MonoDim.Close();

                        Cursor.Current = Cursors.Default;
                    }
                    catch { Cursor.Current = Cursors.Default; }
                }

                return true;
            }
            catch { return false; }
        }
    }
        /// <summary>
        /// bouton d'import des vis
        /// </summary>
        public class Clipper_Import_Fournitures_Divers_Processor : CommandProcessor
    {

        public override bool Execute()
        {

            try
            {
                //creation des logs
                TextWriterTraceListener logFile;
                Import(Context);
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }

        }

        public bool Import(IContext contextlocal)
        {
            try
            {
                using (Clipper_Import_Fournitures_Divers UpdateFourniture = new Clipper_Import_Fournitures_Divers(contextlocal))
                {
                    try
                    {


                        Cursor.Current = Cursors.WaitCursor;


                        UpdateFourniture.Read();
                        UpdateFourniture.Write();
                        UpdateFourniture.Close();

                        Cursor.Current = Cursors.Default;
                    }
                    catch { Cursor.Current = Cursors.Default; }
                }

                return true;
            }
            catch { return false; }
        }


    }
      
    /// </summary>
    ///bouton import des tubes
    public class Clipper_ImportTubes_Processor : CommandProcessor
    {

        Boolean Import_Matiere = true;
        Boolean Tube_Rond = true;
        Boolean Tube_Speciaux = true;
        Boolean Rond = true;
        Boolean Tube_Rectangle = true;
        Boolean Tube_Carre = true;
        Boolean Tube_Flat = true;
        Boolean Fourniture = true;

        /// <summary>
        /// constructeur de l'import des tube, tu stock et des matieres.........
        /// import_matiere,  tube_rond,  rond,  tube_rectangle,  tube_carre, tube_flat,  Tube_Speciaux
        /// </summary>
        /// <param name="import_matiere">true ou false pour declencher l'import</param>
        /// <param name="tube_rond">true ou false pour declencher l'import</param>
        /// <param name="rond">true ou false pour declencher l'import</param>
        /// <param name="tube_rectangle">true ou false pour declencher l'import</param>
        /// <param name="tube_carre">true ou false pour declencher l'import</param>
        /// <param name="tube_flat">true ou false pour declencher l'import</param>
        /// <param name="Tube_Speciaux">true ou false pour declencher l'import</param>
        /// <param name="Fourniture">true ou false pour declencher l'import</param>
        /// 
        public Clipper_ImportTubes_Processor(bool import_matiere, bool tube_rond, bool rond, bool tube_rectangle,  bool tube_carre, bool tube_flat, bool tube_speciaux,bool fourniture)
        {

                                    Import_Matiere = import_matiere;
                                    Tube_Rond = tube_rond;
                                    Rond = rond;
                                    Tube_Rectangle = tube_rectangle;
                                    Tube_Carre = tube_carre;
                                    Tube_Flat = tube_flat;
                                    Tube_Speciaux = tube_speciaux;
                                    Fourniture = fourniture;


        }

        public override bool Execute()
        {

            try
            {
                //creation des logs
                TextWriterTraceListener logFile;



                IContext contextlocal = Context;
                Import(Context);
                return base.Execute();
            }
            catch
            {
                return base.Execute();
            }

        }

        public void Import(IContext contextlocal)
        {
            //creation des logs
            TextWriterTraceListener logFile;
            try
            {   
                //detection du contexte
                Cursor.Current = Cursors.WaitCursor;
                //creation des logs
                //creation du listener
                ////
                if (Import_Matiere) { 
                using (Clipper_Import_Matiere Update_Material = new Clipper_Import_Matiere())
                {
                    try
                    {
                            //Update_Material.Almacam_Update_Material(contextlocal);
                            Update_Material.Import(contextlocal);
                        //Update_Material.Close();
                        Cursor.Current = Cursors.Default;
                    }
                    catch { Cursor.Current = Cursors.Default; }
                }

                }


                if (Tube_Speciaux)
                {
                    using (Clipper_Import_Tube_Speciaux tubesronds = new Clipper_Import_Tube_Speciaux(contextlocal))
                    {
                        try
                        {

                            tubesronds.ReadTubes();
                            tubesronds.WriteTubes();
                            tubesronds.Close();

                        }
                        catch (Exception ie) {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                }



                if (Tube_Rond) { 
                        using (Clipper_Import_Tube_Rond tubesronds = new Clipper_Import_Tube_Rond(contextlocal))
                        {
                            try
                            {

                                tubesronds.ReadTubes();
                                tubesronds.WriteTubes();
                                tubesronds.Close();

                            }
                            catch (Exception ie) {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    
                        }
                }

                if (Rond) { 
                        using (Clipper_Import_Rond tubesronds = new Clipper_Import_Rond(contextlocal))
                        {
                            try
                            {


                                tubesronds.ReadTubes();
                                tubesronds.WriteTubes();
                                tubesronds.Close();

                            }
                            catch (Exception ie) {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                        }
                }

                if (Tube_Rectangle) { 

                        using (Clipper_Import_Tube_Rectangle tubesrectangle = new Clipper_Import_Tube_Rectangle(contextlocal))
                        {
                            try
                            {
                                tubesrectangle.ReadTubes();
                                tubesrectangle.WriteTubes();
                                tubesrectangle.Close();

                            }
                            catch (Exception ie) {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                        }
                }
                if (Tube_Carre) {
                        using (Clipper_Import_Tube_Carre tubescar = new Clipper_Import_Tube_Carre(contextlocal))
                        {
                            try
                            {
                                tubescar.ReadTubes();
                                tubescar.WriteTubes();
                                tubescar.Close();

                            }
                            catch (Exception ie) {
                            MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                        }
                }

                if (Tube_Flat) { 
                    using (Clipper_Import_Tube_Flat  tubesflats = new Clipper_Import_Tube_Flat (contextlocal))
                        {
                            try
                            {
                                tubesflats.ReadTubes();
                                tubesflats.WriteTubes();
                                tubesflats.Close();

                            }
                            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); }

                       
                    }
                }

                
            }
            catch
            {
               
            }

        }
    }
//////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    #region enum, class objets

    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum TypeTube
    {
        Aucun = 0,
        Rond = 4,
        Round = 5,
        Rectangle = 10,
        Flat = 3,
        Speciaux = 6
    }

    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum SheetType
    {      
        ToleNeuve = 0,
        Chute = 1,
       
     }

    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// routine d'import satndards, matiere, tubes.. toles, vis......
    /// </summary>
    //section des commandes
    public enum UniteGestion
    {
        u = 1,
        le_dix = 10,
        le_cent = 100,

    }

    #endregion
 //////////////////////////////////////////////////////////////////////////////////////////
    ///json   
    #region json
    /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <summary>
    /// gestion des fichier json (partira dans import tools plus tard
    /// </summary>
    //section JSON
    public class JsonTools
        {
            //recupee le fichier json dans le dossier de la dll
            private string PATH = Directory.GetCurrentDirectory()+"\\"+ Properties.Resources.jsonimportclipper;//"import-clipper.json";

            public string getJsonStringParametres(string ParameterName)
            {
                
            try {
                        string returnvalue = "";
                        using (StreamReader file = File.OpenText(this.PATH))
                        {
                   
                   
                            using (JsonTextReader reader = new JsonTextReader(file))
                            {
                                JObject o = (JObject)JToken.ReadFrom(reader);
                                returnvalue = (string)o.SelectToken(ParameterName);
                            }
                    

                        }

                        return returnvalue;
                }


            catch (Exception ie) {

                MessageBox.Show(ie.Message,"JsonTools Class" ,MessageBoxButtons.OK, MessageBoxIcon.Error );
                return null;


            }

        }


        }
    #endregion

//////////////////////////////////////////////////////////////////////////////////////////
 /// <summary>
 ///   class Clipper_Material.. class de stockage des matieres
 ///   les codes articles mutlidim sont recuperés dans le champs code article matiere
 /// </summary>
    public class Clipper_Material 
        {
        //initialisation des parametres
        //definiton de l'objetc material
            public string Coarti = "UNDEF"; //code etat (nuance epaisseur)
            public double Prixart=0;// prixart
            public double Densite=1;//
            public string Nuance= "UNDEF";
            public string Etat = "";//code nuance
            public int Type = 0;
            public bool IsMultiDim = false;
            public string Quality; //quality
            public double Thickness = 0;  //DIM1
            public string Materialname = "UNDEF";
            public string Comments = "";
            public string Uidkey;

        /// <summary>
        /// retourne un nom normalisé pour la nuance Nuance*Etat;
        /// </summary>
        /// <returns></returns>
        public string getNuance() { return Nuance + "*" + Etat; }
        /// <summary>
        /// reourne l'epaissseur matiere
        /// </summary>
        /// <returns></returns>
        public double getThickess() { return Thickness; }

        //construit le nom matiere en concatenant nuance*etat
        /// <summary>
        /// update material name
        /// </summary>
        public void setMaterialName()
            {   //setuid de l'objet
                Uidkey = getMaterial_Uidkey(getQualityName(Nuance, Etat), Thickness,Type);
                if (Etat != string.Empty)
                {
                            Quality = Nuance + "*" + Etat;
                            Materialname = Nuance + "*" + Etat + " " + Thickness + " mm";
                }
                else
                {
                        Quality = Nuance ;
                        Materialname = Nuance + " " + Thickness + " mm";
                }
            
            //cle unique pour eviter les doublons
            //Uidkey = Nuance.Trim() + "*" + Etat.Trim() + " " + Thickness;
            

            // return Materialname ;
        }
        /// <summary>
        /// creer l'uidmatiere interne a la dll (structutrée) quality*thickness
        /// </summary>
        /// <param name="qualityname">nuance*etant generalement</param>
        /// <param name="Thickness">epaisseure</param>
        /// <returns></returns>
        public static string getMaterial_Uidkey(string qualityname, double Thickness,int type)
        {
            string uidkey;
            
             uidkey = qualityname + "*" + Thickness.ToString()+"_"+type;

            return uidkey;

            // return Materialname ;
        }
        // <summary>
        /// recupere le nomde la qualité Nuance*Etat*epaisseur
        /// </summary>
        /// <returns></returns>
        public string getQualityName(string Nuance,string Etat) {

            string qualityname;
            //cle unique pour eviter les doublons non stockée
            if (Etat != string.Empty)
            { qualityname = Nuance.Trim() + "*" + Etat.Trim(); }
            else
            { qualityname = Nuance.Trim(); }

            return qualityname;
        }
        /// <summary>
        /// recupere le nom matiere Nuance*Etat*epaisseur
        /// </summary>
        /// <returns></returns>
        public string getMaterialName() {return  Materialname; }

        }
    /// <summary>
    ///   class Clipper_IMport matiere.. derive de clipper article
    /// </summary>
    public class Clipper_Import_Matiere : IDisposable
        {
            private string DSN;
            private string SQL;
            private string SQL_NUANCE_ETAT;
            private JsonTools JSTOOLS;
            private OdbcDataReader TABLE_ARTICLEM_TOLE;
            private OdbcConnection DbConnection;
            private OdbcCommand DbCommand;
            IList<Clipper_Material> CLIPPER_MATERIAL_LIST = new List<Clipper_Material>();
            //recuperation des parametres
            //creation du listener
           TextWriterTraceListener logFile = new TextWriterTraceListener(System.IO.Path.GetTempPath() +"\\"+ Properties.Resources.ImportMatiereLog);
           public IList<Clipper_Material> GetMaterial_List() { return CLIPPER_MATERIAL_LIST; }

            public void Dispose()
            {
                //Dispose(true);
                GC.SuppressFinalize(this);
            }
        //constructeur
        #region constructeur
        public Clipper_Import_Matiere()
            {
                logFile.Write("debut import matiere ");
                Init();
            }
            public Clipper_Import_Matiere(string dsn, string sqlTole)
            {
            //material matiere;                       
            DSN = dsn;
            SQL = sqlTole;
            Init();
            }
        #endregion
        private void Init()
            {    ///premier paramertre dsn = clipper dsn= data source name
                    //var DbConnection =new OdbcConnection("DSN=" + DSN);
            try
                {
                   logFile.WriteLine("initialisatin de la methode");
                    //lecture du fichier json
                    logFile.WriteLine("lecture du fichier Json");
                    // Material matiere = new material(); //
                    JSTOOLS = new JsonTools();
                    //recup parametre dsn //
                    /************/
                    //le dsn est le nom de la connextion odbc paramétre par l'installateur clipper
                    //cette info est paramétrable dans le fichier json
                    DSN = JSTOOLS.getJsonStringParametres("dsn");
                    /************/
                    logFile.WriteLine("DSN demandé"+ DSN);
              
                if (AlmaCamTool.Is_Odbc_Exists(DSN))
                    
                {
                    DbConnection = new OdbcConnection("DSN=" + DSN);                    
                    DbConnection.Open();
                    DbCommand = DbConnection.CreateCommand();
                    logFile.WriteLine("etat de la connexion " + DbConnection.State.ToString());
                    Get_Clippper_Material();
                }
                else{ throw new Missing_Obdc_Exception(DSN); }

     


            }

            catch (Exception ie)
                {
                logFile.WriteLine("ERREUR : " + ie.Message.ToString());
                }
            finally
            {
                //DbConnection.Close();
                this.Dispose();
            }


            }

        /// <summary>
        ///  recupere dans une table CLIPPER_MATERIAL_LIST
        /// </summary>
        private void Get_Clippper_Material()
            {
                //IList <Material> materialist = new List<Material>();
                //string prefix; // prefixe article matiere
                //string grade;//code nuance
                //string name; //code etat (nuance epaisseur)
                //double price; // prix
                //double densité;
                //double Thickness;
                int ii = 0;
            logFile.WriteLine("paramétrage de la requete du fichier json " );
            logFile.WriteLine("monodim = " + JSTOOLS.getJsonStringParametres("multidim"));
            //recuperation des nuance etat--> 304L*DKP
            //on recherche toutes les matieres mono et multidim

           
            this.SQL_NUANCE_ETAT = JSTOOLS.getJsonStringParametres("sql.nuance_etat");
                                                                                      //coupel nuance etat
                                                                                      //Material matiere = new material();
                                                                                      //getJsonParametres();
            this.DbCommand.CommandText = this.SQL_NUANCE_ETAT;
            logFile.WriteLine("requete utilisée = " + this.SQL_NUANCE_ETAT);
            // requete type matiere recuperer dans la section nuance_etat //
            TABLE_ARTICLEM_TOLE = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM_TOLE.Read())
                {
                    ii++;
                    Clipper_Material material = new Clipper_Material();
                    material.Nuance = TABLE_ARTICLEM_TOLE["CODENUANCE"].ToString().Trim();
                    material.Etat = TABLE_ARTICLEM_TOLE["CODEETAT"].ToString().Trim();
                    material.Prixart = Convert.ToDouble(TABLE_ARTICLEM_TOLE["PRIXART"]);
                    material.Densite = Convert.ToDouble(TABLE_ARTICLEM_TOLE["DENSITE"]);
                    material.Type = Convert.ToInt32(TABLE_ARTICLEM_TOLE["TYPE"]);
                    material.IsMultiDim = Convert.ToBoolean(TABLE_ARTICLEM_TOLE["MULTIDIM"]);
                
                //attention pour les tole monodim l'epaisseur est stock en dim3.
                //recuperation des codes multidim par defaut

                if (material.IsMultiDim) {
                    material.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TOLE["EPAISSEUR"]);
                    material.Coarti = TABLE_ARTICLEM_TOLE["COARTI"].ToString().Trim();
                }
                else {
                    material.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TOLE["DIM3"]);
                    material.Coarti = string.Empty;
                }


                //on voit dans la liste le code article
                material.Comments ="" ;//TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim();
                //issue du tube ou non
                if (TABLE_ARTICLEM_TOLE["TYPE"].ToString().Trim() == "3")

                {   //plats
                    material.Comments = TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim();
                    material.setMaterialName();
                }
                else {
                    //tube et autres
                    material.Comments = "Matiere_Tube " + TABLE_ARTICLEM_TOLE["MATERIAL_COMMENTS"].ToString().Trim(); }
                    material.setMaterialName();
                //teste d'integrite sur la matiere
               if (CheckDataintegrity(material))

                {
                    // on controle les doublons ici //
                    if (CLIPPER_MATERIAL_LIST.Where(m => m.Uidkey == material.Uidkey).Count() == 0)
                    { CLIPPER_MATERIAL_LIST.Add(material); }
                    else
                    {
                        logFile.WriteLine("doublon detecter sur la matiere, la matiere ne sera pas ajouté " + material.Uidkey);
                    }
                }

            };
            logFile.WriteLine("reconstitution de la table des matieres terminés : " + this.CLIPPER_MATERIAL_LIST.Count().ToString()+ " trouvées ");
          

        }
        /// <summary>
        /// fermeture de la connexion odbc
        /// </summary>

        public void Close()
            {

                // reader.Close();

                DSN = "";
                SQL = "";

                if (TABLE_ARTICLEM_TOLE != null)
                {
                    TABLE_ARTICLEM_TOLE.Close();

                }

                DbConnection.Dispose();
                DbCommand.Dispose();
                CLIPPER_MATERIAL_LIST = null;
                DbCommand.Dispose();
                DbConnection.Close();

                logFile.Close();

        }
        
        /// <summary>
        /// phase d'import
        /// ecrit les nouvelles matieres dans almacam
        /// </summary>
        /// <param name="contextlocal"></param>
        
        public void Import(IContext contextlocal)
        {
            //  logFile.WriteLine("mise a jour des matieres dans la base " + contextlocal.Model.DatabaseName);
            //ecrirture de la liste des matiere//
            //recuperation de la liste des matiere almacam
                IEntityList qualityentitylist = contextlocal.EntityManager.GetEntityList("_QUALITY");
                IList<string> qualities_To_Create = new List<string>();
                qualityentitylist.Fill(false);

                IList<IEntity> qualitylist = new List<IEntity>();

                //select distincte supprime les doublons
                qualitylist = qualityentitylist.Distinct().ToList();
                bool newdatabase = qualitylist.Count == 0;
                ////liste des qualité de materiaux
                IDictionary<long,string> updatedstringqualitylist = new Dictionary<long,string>();

            //////
            ///////detection des qualités a creer//// 
            logFile.WriteLine("creation dela liste des qualités de clipper ");
            foreach (Clipper_Material m in CLIPPER_MATERIAL_LIST)
            {
          
                if (qualities_To_Create.Contains(m.Quality)==false)
                {  qualities_To_Create.Add(m.Quality);  }
               

            }
            logFile.WriteLine(qualities_To_Create.Count().ToString()+" qualités trouvées ");
            
            ///////////
            ///creation des qualité dans la base si necessaire
                        
            foreach (string quality in qualities_To_Create)
            {

                // Find material           
                IEntity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                if (currentQuality == null)
                { 
                    //creation de la nuance et sauvegarde// ou retour de la qualité courente
                    currentQuality = contextlocal.EntityManager.CreateEntity("_QUALITY");
                    currentQuality.SetFieldValue("_NAME", quality);
                    logFile.WriteLine("creation de la nouvelle qualité  "+ quality);
                    currentQuality.Save();

                }




            }

            logFile.WriteLine( " qualités crées ");
            qualityentitylist.Fill(false);
            qualitylist = qualityentitylist.ToList();

            ///////////
            ///creation des matieres assocées au qualités
            //mise a jour de la liste des matiere
            // on ne prends que les matiere de type 3  (tole/plat)         

            foreach (Clipper_Material m in CLIPPER_MATERIAL_LIST) {

                // Find material           
                IEntity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(m.Quality)).FirstOrDefault();
                //update
                currentQuality.SetFieldValue("_NAME", m.Quality);
                currentQuality.SetFieldValue("_DENSITY", (m.Densite)/1000);
                //calcul du prix moyen toutes epaisseurs confondues
                currentQuality.SetFieldValue("_BUY_COST",( CLIPPER_MATERIAL_LIST.Average(p => p.Prixart)));
                currentQuality.SetFieldValue("_OFFCUT_COST", (CLIPPER_MATERIAL_LIST.Average(p => p.Prixart)/1000));
                currentQuality.SetFieldValue("_COMMENTS", m.Comments); 
                currentQuality.Save();
                //on rempli la liste des qualités
                if (currentQuality != null) { 
                        if (!updatedstringqualitylist.ContainsKey(currentQuality.Id) )
                        {
                            updatedstringqualitylist.Add(currentQuality.Id,m.Quality);
                        }
                }

            }


            //création des matieres   ---> pas d'unicité (voir dans la requete)
            //on rempli une liste de keypair de type uidkey,ientity matiere pour detecter les doublons uidkey est une cle interne a l'obkey clipper matierial
            //
            logFile.WriteLine("creation des matieres assocées aux qualités  ");
            foreach (var currentstringquality in updatedstringqualitylist)
            {
                {   //recuperation de la lisre de matieres associée a la qualité choisie
                                     
                    IEntityList almacam_materialentitylist = contextlocal.EntityManager.GetEntityList("_MATERIAL", "_QUALITY", ConditionOperator.Equal, currentstringquality.Key);
                    almacam_materialentitylist.Fill(false);

                    //construction de la liste des keypair (voir pour mettre un distionnaire si trop lent)
                    IList<IEntity> almacam_material_list = new List<IEntity>();
                    almacam_material_list = almacam_materialentitylist.ToList();
                    //detection des doublons
                    List<KeyValuePair<string,IEntity >> almacam_material_UiKey = new List<KeyValuePair<string,IEntity>>();                   
                    foreach (Entity e in almacam_material_list)
                    {
                        string quality = e.GetFieldValueAsEntity("_QUALITY").GetFieldValueAsString("_NAME");
                        double thickness = e.GetFieldValueAsDouble("_THICKNESS");
                        //Clipper_Material.getMaterial_Uidkey(quality, thickness);
                        almacam_material_UiKey.Add(new KeyValuePair<string,IEntity>(quality +"*"+thickness,e));
                    }

                    //on se limite aux matiere des plats/toles
                    foreach (Clipper_Material m in this.CLIPPER_MATERIAL_LIST.Where(t => t.Type == 3))
                    {
                        if (m.Quality == currentstringquality.Value && m.Type==3)
                        {

                            IEntity currentmaterial = null;

                            //IEntity currentmaterial = almacam_material_list.Where(q => q.GetFieldValueAsString("_NAME").Equals(m.getMaterialName())).FirstOrDefault();
                            //si la clé est retrouvé alors c'est un update sinon c'est une creation
                            if (almacam_material_UiKey.Where(kvp => kvp.Key== m.Uidkey.Split('_')[0]).Count()>0)
                            {
                                currentmaterial = almacam_material_UiKey.First(kvp => kvp.Key == m.Uidkey.Split('_')[0]).Value;
                            }
                           

                            //creation
                            if (currentmaterial == null)
                            { 
                                //creation de la matiere et sauvegarde//
                                currentmaterial = contextlocal.EntityManager.CreateEntity("_MATERIAL");
                                logFile.WriteLine("creation de la matieres   " + m.getMaterialName());
                                currentmaterial.Save();

                            }

                            //update and save
                                logFile.WriteLine("mise à jour de la matiere :   " + m.getMaterialName());
                           
                                currentmaterial.SetFieldValue("_NAME", m.getMaterialName());
                                currentmaterial.SetFieldValue("_QUALITY", currentstringquality.Key);
                                currentmaterial.SetFieldValue("_THICKNESS", m.Thickness);
                                currentmaterial.SetFieldValue("_BUY_COST", (m.Prixart)/1000);
                                currentmaterial.SetFieldValue("_CLIPPER_CODE_ARTICLE", m.Coarti);
                                currentmaterial.SetFieldValue("_COMMENTS",m.Comments);
                                
                            currentmaterial.Save();
                            //on set le nom standard
                            CommonModelBuilder.ComputeMaterialName(currentmaterial.Context, currentmaterial.GetFieldValueAsEntity("_QUALITY"), currentmaterial);
                            currentmaterial.Save();
                        }
                    }

                    //pas utils mais au cas ou
                    almacam_materialentitylist = null;
                    almacam_material_list = null;


                }
            



            }
            //*//

            logFile.Flush();

            this.Close();
        }
        public bool CheckDataintegrity(Clipper_Material material)
            {   //epaisseur max
                //si l'article existe deja on le 'import pas
                bool integrity = true;
            //cette matiere existe deja selon les clitter clippers
            //if (CLIPPER_MATERIAL_LIST.Select(m => m.Materialname == material.Materialname).Count() > 0)
            //foreach (Clipper_Material_)
            //{ integrity = false; }
          
            return integrity;
        }




    }
    /// <summary>
    /// 
    /// 
    /// 
    /// 
    /// 
    /// 
    


    /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    /// </summary>
    /// concerne l'import des articles
    /// 
    /// <summary>
    ///   classe de base clipper article
    /// </summary>
    public class Clipper_Article
    {
        //string prefix; // prefixe article matiere
        //creation du listener

        public IContext contextlocal;
        public long AMCLEUNIK;
        public string InternalName; // nom interne
        //fournisseur

        public string FOURN;
        public string UniteGest = "u"; //unite de gestion : u ou le cent , le dix...
        public string UnitePrix = "u";
        public string COARTI = "UNDEF"; //code etat (nuance epaisseur)
        public double PRIXART;// prixart
        public double Densite;//
        public IEntity Material;
        public string Nuance;
        public string Etat = "";

        //codification
        public string Name = "UNDEF";
        public string COFA = "UNDEF";
        public string DESA1 = "";
        public string DESA2 = "";
        private string DSN;
        //dimensions
        //public Dictionary<string, double> DIM;
        /// <summary>
        /// conexion
        /// </summary>
         public OdbcDataReader TABLE_ARTICLEM;//TABLE_ARTICLEM;
         public OdbcConnection DbConnection;
         public OdbcCommand DbCommand;
         public JsonTools JSTOOLS;

        public void Dispose()
            {
            TABLE_ARTICLEM = null;
            //TABLE_ARTICLEM_ROND = null;
            DSN = null;
            }
            
        IList<Clipper_Material> TUBE_LIST = new List<Clipper_Material>();
        //IList<string> Articles = new List<string>();
        //IList<Clipper_Material> SHEET_LIST = new List<Clipper_Material>();
        /// <summary>
        /// constructeur
        /// </summary>
        public Clipper_Article()
        {
         // DIM = new Dictionary<string, double>();

        }
        /// <summary>
        /// lecture des donnees cam et clip
        /// </summary>
        public virtual void Read()
        {

            
        }
        /// <summary>
        /// ecriture dans cam et eventuelement maj
        /// </summary>
        public virtual void Write()
        {

           
        }
        /// <summary>
        /// maj dans cam 
        /// </summary>
        public virtual void Update()
        {

            
        }

        /// <summary>
        /// creation de la connexion odbc
        /// </summary>
        public void Odbc_Connexion()
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {

                //Material matiere = new material();
                JSTOOLS = new JsonTools();
                //recup parametre dsn
                DSN = JSTOOLS.getJsonStringParametres("dsn");

                if (AlmaCamTool.Is_Odbc_Exists(DSN)) { 

                        //recupe requete
                        DbConnection = null;
                        DbCommand = null;
                        this.DbConnection = new OdbcConnection("DSN=" + DSN);
                        this.DbConnection.Open();
                        this.DbCommand = DbConnection.CreateCommand();
                        //contextlocal.TraceLogger.TraceInformation("Creation du lien ODBC..");
                            //logFile.WriteLine("connexion ok    : " + DSN);
                }

            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode ,  ex.Message);
                Environment.Exit(0);
            }
        }
        //permet d'ecrire dans le log almacam
        public void Odbc_Connexion(IContext contextlocal)
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {

                //Material matiere = new material();
                JSTOOLS = new JsonTools();
                //recup parametre dsn
                DSN = JSTOOLS.getJsonStringParametres("dsn");

                if (AlmaCamTool.Is_Odbc_Exists(DSN))
                {

                    //recupe requete
                    DbConnection = null;
                    DbCommand = null;
                    this.DbConnection = new OdbcConnection("DSN=" + DSN);
                    this.DbConnection.Open();
                    this.DbCommand = DbConnection.CreateCommand();
                    contextlocal.TraceLogger.TraceInformation("Creation du lien ODBC..");
                    //logFile.WriteLine("connexion ok    : " + DSN);
                }

            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                Environment.Exit(0);
            }
        }
        /// <summary>
        /// creation de la connexion ole
        /// </summary>
        public void Ole_Connexion()
        {    ///premier paramertre dsn = clipper dsn= data source name
            try
            {

               


            }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                Environment.Exit(0);
            }
        }
        /// <summary>
        /// recuperation ou conversion numeric de donnees sql, retourn 0 si texte vide
        /// </summary>
        /// <param name="fieldname"></param>
        /// <returns></returns>
        public string getSqlNumericValue(string fieldname)
        {
            try
            {
                //string value;
                if (TABLE_ARTICLEM[fieldname].ToString().Trim() == string.Empty) { return "0"; }
                else { return TABLE_ARTICLEM[fieldname].ToString().Trim(); }

            }
            // catch { return "0"; }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message   + fieldname + "Not found ");
                return "0";
               
            }

        }

        /// <summary>
        /// recuperation ou conversion numeric de donnees sql, retourn 0 si texte vide
        /// </summary>
        /// <param name="fieldname"></param>
        /// <returns></returns>
        public int getSqlIntValue(string fieldname)
        {
            try
            {
                //string value;
                if (TABLE_ARTICLEM[fieldname].ToString().Trim() == string.Empty) { return 0; }
                else {
                    string temp =  TABLE_ARTICLEM[fieldname].ToString().Trim();
                    return Convert.ToInt32(temp);

                }

            }
            // catch { return "0"; }
            catch (Exception ex)
            {
                string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
                MessageBox.Show(methode, ex.Message);
                return 0;

            }

        }

        public void Close()
        {
            
            DSN = "";
            //SQL = "";
            TABLE_ARTICLEM.Close();
            DbConnection.Dispose();
            DbCommand.Dispose();
            DbCommand.Dispose();
            DbConnection.Close();
           
            //logFile.WriteLine("connexion    : Fin" );
            //logFile.Close();
            //logFile.Flush();
            //logFile.Dispose();


        }

        #region Matiere
        public IEntity GetGrade(IContext contextlocal, string nuance, string etat)
        {
            try
            {
                //en cas de lenteru on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons
                IEntityList grades;
                IEntity grade;
                string qualityname;
                if (etat == "") { qualityname = nuance; } else { qualityname = nuance + "*" + etat; }

                grades = contextlocal.EntityManager.GetEntityList("_QUALITY", "_NAME", ConditionOperator.Equal, qualityname);
                grades.Fill(false);

                if (grades.Count() > 0)
                {
                    grade = grades.FirstOrDefault();
                

                }
                else
                {
                    grade = null;
                  
                }


                return grade;


            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return null; }
        }

        public double GetPrice(string unitegestion, double price)
        {
            try
            {

                switch (unitegestion)
                {
                    case "le cent":
                        price = price / 100;
                        break;

                    case "le dix":
                        price = price / 10;
                        break;
                    default:
                        //price = price;
                        break;
                }


                return price;


            }
            catch { return 0; }
        }

    }



    #endregion

    #region class des fourniture (vis, ecrous.... cartons)
   
    /// <summary>
    /// propriété specifiques founitures divers
    /// </summary>
    public class Founiture_Divers : Clipper_Article, IDisposable
    {
      
    }

    #endregion


    #region class des tole (vis, ecrous.... cartons)
    /// <summary>
    /// recuperation des toles monodims
    /// </summary>
    public class Clipper_Article_Toles_MonoDim : Clipper_Article, IDisposable
    {
     
        private string entityType = "_SHEET";
        public List<Flat> Clipper_articles_Stock;
        //public Dictionary<string, IEntity> Almacam_Stock;
        public Dictionary<string, IEntity> Almacam_Stock;
        /* public JsonTools JSTOOLS; */
        //IList<Clipper_Material> TOLES_MONO_DIM_LIST = new List<Clipper_Material>();

        public Clipper_Article_Toles_MonoDim(IContext context)
        {
            this.contextlocal = context;
            Odbc_Connexion();
            //getAlmacam_Existing_Format_List(entityType);
            Almacam_Stock=getAlmacam_Stock_Format_List();
        }


        TextWriterTraceListener logFilesheetMonodim = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "SHEET_MONO_" + Properties.Resources.Import_Sheet_Mono);



        /// <summary>
        /// retourn une liste d'exclusion des formats
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> getAlmacam_Existing_Format_List(string entitype)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();

               
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetFieldValueAsString("_REFERENCE").Trim();
                    if (uid != string.Empty)
                    {
                        if (exclusion.Contains(uid) == false  )
                        { exclusion.Add(uid); }
                    }
                    

                }

                return exclusion;
            }

            catch
            {

                return null;
            }


        }

        /// <summary>
        /// retourne la liste des toles et leur format
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns>dicionnaire string entity sheet </returns>
        public Dictionary<string,IEntity> getAlmacam_Stock_Format_List()
        {
            try
            {
                string entitype = "_STOCK";
                Dictionary<string, IEntity> sheetlist = new Dictionary<string, IEntity>();

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
               //List<string> exclusion = new List<string>();

                //recuperation du format de chaque entity de stock
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetFieldValueAsString("_NAME").Trim();
                    IEntity e = ex.GetFieldInternalValueAsEntity("_SHEET");
                    uid += "_"+e.Id;
                    if (uid != string.Empty)
                    {
                        IEntity rst = null;
                        sheetlist.TryGetValue(uid, out rst);
                        if (rst==null)
                        { sheetlist.Add(uid,e); }
                    }

                    
                }



                // Create the query.
                // The first line could also be written as "var studentQuery ="



                
                
                return sheetlist;





            }

            catch
            {

                return null;
            }


        }


        /// <summary>
        /// rempli la table Clipper_articles_Stock
        /// pour le moment seules les toles neuves et format commerciaux sont recupere dans cette list
        /// object tole article
        ///flat.COARTI 
        ///flat.Nuance
        ///flat.Etat
        ///flat.PRIXART
        ///bool multidim
        ///flat.MultiDim
        //recuperation des dimenssions
        ///flat.Longueur
        ///flat.Largeur
        ///flat.Epaisseur
        ///flat.Densite
        ///flat.COFA
        ///flat.InternalName = nuance_etat + "*" + flat.Longueur + "*" + flat.Largeur + "*" + flat.Epaisseur;
        ///flat.type= SheetType.Chute;
        ///flat.type = SheetType.ToleNeuve; //forcement tole neuve
        ///flat.commercial = true;
        ///
        /// </summary>
        public override void  Read() {


            Clipper_articles_Stock = new List<Flat>();


            try
            {
                
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.All_Flat");
                //logFileTubeRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    Flat flat = new Flat();
                    flat.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();

                    flat.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    flat.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    string nuance_etat= flat.Etat == "" ? flat.Nuance  : flat.Nuance + "*" + flat.Etat;

                    flat.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    bool multidim = TABLE_ARTICLEM["MULTIDIM"].ToString() == "0" ? false : true;
                    flat.MultiDim = multidim;
                    //recuperation des dimenssions
                    if (flat.MultiDim)
                    {
                        flat.Longueur = Convert.ToDouble(getSqlNumericValue("LOMULTIDIM"));
                        flat.Largeur = Convert.ToDouble(getSqlNumericValue("LAMULTIDIM"));
                        flat.Epaisseur = Convert.ToDouble(getSqlNumericValue("EPMULTIDIM"));
                    }
                    else
                    {
                        flat.Longueur = Convert.ToDouble(getSqlNumericValue("LOMONODIM"));
                        flat.Largeur = Convert.ToDouble(getSqlNumericValue("LAMONODIM"));
                        flat.Epaisseur = Convert.ToDouble(getSqlNumericValue("EPMONODIM"));
                    }

                     flat.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    flat.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim();

                    flat.InternalName = nuance_etat + "*" + flat.Longueur + "*" + flat.Largeur + "*" + flat.Epaisseur;

                    //int type = TABLE_ARTICLEM["CHUTE"].ToString() == "N" ? 0 : 1;
                    if (TABLE_ARTICLEM["CHUTE"].ToString() == "O") {
                        flat.type= SheetType.Chute;
                    }
                    else if(TABLE_ARTICLEM["CHUTE"].ToString() == "1") {
                        flat.type = SheetType.ToleNeuve; //forcement tole neuve
                        flat.commercial = true;
                    }
                    else
                    {
                        flat.type = SheetType.ToleNeuve;
                    }

                    //condition d'ajout dans la liste ==> format commercial multidim
                    //format monodim  ce qui assure un code article unique
                    if(flat.MultiDim=false || flat.commercial)
                    {
                        this.Clipper_articles_Stock.Add(flat);

                    }
                    logFilesheetMonodim.Write("sheet : " + flat.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
             
            }

            catch (Exception ie)
            {
                Clipper_articles_Stock = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilesheetMonodim.Close();
                
            }

















        }
        /// <summary>
        /// ecriture des tole dans almacam
        /// </summary>
        /// 

        public override void Write()
        {


            try
            {

                UpdatePrice();

            }

            catch (Exception ie)
            {
                Clipper_articles_Stock = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilesheetMonodim.Close();

            }



        }
        public  void UpdatePrice() {

             try {
                //this.Clipper_articles_Stock; //liste du stock clip
                //this.Almacam_Stock; liste du stock cam
               if (Clipper_articles_Stock.Count>0 & Clipper_articles_Stock.Count()>0)
                {
                    foreach (Flat flat in Clipper_articles_Stock)
                    {
                        var sheetlist = new Dictionary<string, IEntity>();

                        sheetlist = Almacam_Stock.Where(p => p.Key.Contains(flat.COARTI + "_")).ToDictionary(p => p.Key, p => p.Value);
                        foreach (IEntity e in sheetlist.Values)
                        {
                            double poids = e.GetFieldValueAsDouble("_WEIGHT") / 1000;
                            e.SetFieldValue("_BUY_COST", flat.PRIXART * poids);
                            e.SetFieldValue("_COMMENTS", "Prix = " + flat.PRIXART + " €/kg");
                            e.SetFieldValue("_AS_SPECIFIC_COST", true);
                            e.Save();
                        }

                        sheetlist = null;
                    }
                    //        flat.COARTI
                    //srtring coda = Clipper_articles_Stock[flat.COARTI]


                } 




            }
            catch {


            }
            



        }
        //public virtual void DimInterpretor() { }
        public virtual void Close() { Clipper_articles_Stock = null; }
    }
    /// <summary>
    /// recuperation des toles multidim
    /// </summary>
    public class Flat : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        public SheetType type; // = SheetType.ToleNeuve;
        public bool MultiDim = false;
        public bool commercial = false;
        public double Longueur = 0;
        public double Largeur = 0;
        public double Epaisseur = 0;


        

    }
    #endregion



    #region Class des Tubes

    /// <summary>
    ///   class Clipper_Article_Tube.. derive de clipper article
    /// </summary>
    public class Clipper_Article_Tube :Clipper_Article, IDisposable
    {
        //string prefix; // prefixe article matiere
        //creation du listener
        //tools
        public string Section_Key; //"_SECTION_CIRCLE" or....
        public List<string> SectionExclusion = new List<string>();
        public List<string> TubeExclusion = new List<string>();
       
        
      

       /* public JsonTools JSTOOLS;*/

        IList<Clipper_Material> TUBE_LIST = new List<Clipper_Material>();

        public Clipper_Article_Tube()
        {

            Odbc_Connexion();
        }



        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> getSectionExclusionList(string entitype)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();

                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetImplementEntity("_SECTION").GetFieldValueAsString("_NAME").Trim();
                    ex.GetImplementEntity(Section_Key);
                    if (exclusion.Contains(uid) == false)
                    { exclusion.Add(uid); }

                }


                return exclusion;
            }

            catch
            {

                return null;
            }


        }

        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public List<string> getExclusionList(string entitype, string uid_field)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                List<string> exclusion = new List<string>();

                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {
                    string uid = ex.GetFieldValueAsString(uid_field).Trim();

                    ex.GetImplementEntity(Section_Key);
                    if (exclusion.Contains(uid) == false)
                    { exclusion.Add(uid); }

                }


                return exclusion;
            }

            catch
            {

                return null;
            }


        }
       

        public long getExistingSection(string section_name,string section_key)
        {

            try
            {
                //en cas de lenteur on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons

                IExtendedEntityList sections;
                IExtendedEntity section=null;
                long section_id =0;

                if (section_name != "")
                {
                    //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);

                    sections = contextlocal.EntityManager.GetExtendedEntityList(contextlocal.Kernel.GetEntityType(section_key));
                    sections.Fill(false);

                    foreach (IExtendedEntity s in sections)
                    {
                        object o = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION\\_NAME");
                        if (o.ToString() == section_name) {
                            section_id = s.Id32;
                            section = s;
                            break;

                        }
                    }





                    return section_id;
                 
                }
                else
                { return 0; }
             


            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return 0; }


        }
        public long getExistingImplementedSection(string section_name, string section_key)
        {

            try
            {
                //en cas de lenteur on pourrais stocker toutes les gardes poour eviter de faire troo de requetes
                //nous verrons

                IExtendedEntityList sections;
                //IExtendedEntity section = null;
                long section_id = 0;

                if (section_name != "")
                {
                    //sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);

                    sections = contextlocal.EntityManager.GetExtendedEntityList(contextlocal.Kernel.GetEntityType(section_key));
                    sections.Fill(false);

                    foreach (IExtendedEntity s in sections)
                    
                    {
                        object o = s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION\\_NAME");
                        object id= s.GetFieldValue(section_key + "\\IMPLEMENT__SECTION");
                        if (o.ToString() == section_name)
                        {
                            section_id =(long) id;

                            //section = s;
                            break;

                        }
                        
                    }





                    return section_id;

                }
                else
                { return 0; }



            }
            catch (Exception ie) { MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error); return 0; }


        }
        /// <summary>
        /// lecture des tube de la base clipper
        /// </summary>
        public virtual void ReadTubes() { }
        /// <summary>
        /// ecriture des tubes dans almacam
        /// </summary>
        public virtual void WriteTubes() { }

        public virtual void DimInterpretor() { }
    }
    /// <summary>
    /// class TubeRond.. derive de clipper article
    /// class tube rond
    /// on ajoute longeur et diametre
    /// </summary>
    public class TubeRond : Clipper_Article, IDisposable
    {
        

        //private List<TubeRond>listeTubeRond;
        private TypeTube type=TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        public double epaisseur;


    }
    /// <summary>
    /// class TubeRec.. derive de clipper article
    /// </summary>
    public class TubeRec : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Largeur;
        public double Hauteur;
        public double Longueur;
        public double epaisseur;

    }

    public class TubeSpe : Clipper_Article, IDisposable
    {

        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Speciaux;
        public double Longueur;
        public string Section;
        public string Section_Spe_Key;
        
       
    }


    /// <summary>
    /// recuperation des tubes ronds
    /// </summary>

    /// <summary>
    /// recuperation des tubes ronds dans une liste de d'objet tubes ronds
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>
    public class Clipper_Import_Tube_Rond : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        private List<TubeRond> List_Tube_Ronds;

        TextWriterTraceListener logFileTubeRond = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "TUBE_ROND_" + Properties.Resources.ImportTubeLog);

        /// <summary>
        /// constructeur obligatoire pour ajuster les elements a controler
        /// </summary>
        /// <param name="context"></param>
        public Clipper_Import_Tube_Rond(IContext context)
        {

            this.contextlocal = context;
            Section_Key = "_SECTION_CIRCLE";
            SectionExclusion = getSectionExclusionList(Section_Key);
            TubeExclusion = getExclusionList("_BARTUBE", "_REFERENCE");

        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond de la base clipper
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Tube_Ronds = new List<TubeRond>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeRond");
                logFileTubeRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRond tuberond = new TubeRond();
                    tuberond.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberond.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberond.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberond.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TUBE["EPAISSEUR"]);
                    tuberond.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    tuberond.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tuberond.Diametre = Convert.ToDouble(getSqlNumericValue("DIAM"));
                    tuberond.epaisseur = Convert.ToDouble(getSqlNumericValue("EPAISSEUR"));
                    tuberond.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Tube_Ronds.Add(tuberond);
                    logFileTubeRond.Write("Tube : " + tuberond.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Tube_Ronds = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileTubeRond.Close();

                /*return listeTubeRond;*/
            }

        }

        /// <summary>
        /// ecriture dans la base almacam
        /// </summary>
        public override void WriteTubes()
        {


            if (List_Tube_Ronds.Any())
            {

                foreach (TubeRond tuberond in List_Tube_Ronds)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    string key = Guid.NewGuid().ToString();
                    //creation de la section
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);

                    IEntity sectionQuality = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    string name = String.Format("Tube_Rond*{0}*{1}", tuberond.Diametre, tuberond.epaisseur);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);// +"x"+ep.ToString());;


                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                    sectionEntity.SetFieldValue("_P_D", tuberond.Diametre);
                    sectionEntity.SetFieldValue("_P_T", tuberond.epaisseur);
                    //descpription //
                    string description = String.Format("Tube Rond Diamete={0} mm Eaisseur={1} mm", tuberond.Diametre, tuberond.epaisseur);

                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", description);

                    //sectionEntity.SetFieldValue("_P_T", this.ep);           
                    sectionEntity.Complete = true;

                    logFileTubeRond.Write("Tube : " + tuberond.COARTI + " sauvegarde");
                    //creation de la qualité
                    tuberond.Material = tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat);

                    //recuperation de la section
                    if (SectionExclusion.Contains(name) == false)

                    {   //si n existe pas on recupere la nouvelle section //
                        sectionEntity.Save();
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                    }
                    else
                    //si elle existe deja on recupere la section //
                    { sectionQuality.SetFieldValue("_SECTION", getExistingSection(name, Section_Key)); }


                    if (TubeExclusion.Contains(tuberond.COARTI) == false)
                    {
                        sectionQuality.SetFieldValue("_QUALITY", tuberond.Material.Id32);
                        sectionQuality.SetFieldValue("_BUY_COST", tuberond.PRIXART);
                        sectionQuality.Save();
                        logFileTubeRond.Write("section " + description + " sauvegardée");
                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tuberond.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tuberond.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tuberond.Longueur);
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntity.GetImplementEntity("_SECTION").Id); //sectionEntity.GetImplementEntity("_SECTION").Id
                        barreEntity.Save();
                    }
                    logFileTubeRond.Write("section " + description + "lng " + tuberond.Longueur + " sauvegardée");
                }
            }

        }



    }


    /// <summary>
    /// recuperation des ronds dans une liste de d'objet rond
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>
    public class Clipper_Import_Rond : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rond;
        public double Diametre;
        public double Longueur;
        private List<TubeRond> List_Ronds;

        TextWriterTraceListener logFileRond = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_ROND_" + Properties.Resources.ImportTubeLog);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public Clipper_Import_Rond(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_PLAIN_ROUND";
            SectionExclusion = getSectionExclusionList(Section_Key);
            TubeExclusion = getExclusionList("_BARTUBE", "_REFERENCE");

        }


        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Ronds = new List<TubeRond>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.Rond");
                logFileRond.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRond tuberond = new TubeRond();
                    tuberond.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberond.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberond.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberond.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM_TUBE["EPAISSEUR"]);
                    tuberond.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    tuberond.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tuberond.Diametre = Convert.ToDouble(getSqlNumericValue("DIAM"));
                    //tuberond.epaisseur= Convert.ToDouble(getSqlNumericValue("EPAISSEUR"));
                    tuberond.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberond.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Ronds.Add(tuberond);
                    logFileRond.Write("Tube : " + tuberond.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Ronds = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileRond.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        public override void WriteTubes()
        {
            if (List_Ronds.Any())
            {
                foreach (TubeRond tuberond in List_Ronds)
                {
                    //tuberond.create(contextlocal);

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    string key = Guid.NewGuid().ToString();
                    //creation de la section
                    //sectionEntity = contextlocal.EntityManager.CreateEntity("_SECTION_CIRCLE");
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);
                    IEntity sectionQuality = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    string name = String.Format("Rond*{0}", tuberond.Diametre);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);// +"x"+ep.ToString());

                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                    sectionEntity.SetFieldValue("_P_D", tuberond.Diametre);

                    //descpription //

                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", String.Format("Rond Diametre={0} mm", tuberond.Diametre, tuberond.epaisseur));

                    sectionEntity.Complete = true;
                    sectionEntity.Save();
                    logFileRond.Write("Tube : " + tuberond.COARTI + " sauvegarde");
                    //creation de la qualité
                    tuberond.Material = tuberond.GetGrade(contextlocal, tuberond.Nuance, tuberond.Etat);

                    //recuperation de la section
                    if (SectionExclusion.Contains(name) == false)

                    {   //si n existe pas on recupere la nouvelle section //
                        sectionEntity.Save();
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                    }
                    else
                    //si elle existe deja on recupere la section //
                    { sectionQuality.SetFieldValue("_SECTION", getExistingSection(name, Section_Key)); }

                    if (TubeExclusion.Contains(tuberond.COARTI) == false)
                    {
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                        sectionQuality.SetFieldValue("_QUALITY", tuberond.Material.Id32);
                        sectionQuality.SetFieldValue("_BUY_COST", tuberond.PRIXART);

                        sectionQuality.Save();
                        logFileRond.Write("section " + name + " sauvegardée");
                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tuberond.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tuberond.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tuberond.Longueur);
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntity.GetImplementEntity("_SECTION").Id); //sectionEntity.GetImplementEntity("_SECTION").Id
                        barreEntity.Save();
                        logFileRond.Write("section " + name + "lng " + tuberond.Longueur + " sauvegardée");
                    }
                }
            }

        }




    }


    /// <summary>
    /// recuperation des rectangles dans une liste de d'objet rectangles
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>

    public class Clipper_Import_Tube_Rectangle : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Diametre;
        public double Longueur;
        private List<TubeRec> List_Recs;

        TextWriterTraceListener logFileRec = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_REC_" + Properties.Resources.ImportTubeLog);


        public Clipper_Import_Tube_Rectangle(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_SHARP_RECTANGLE";
            SectionExclusion = getSectionExclusionList(Section_Key);
            TubeExclusion = getExclusionList("_BARTUBE", "_REFERENCE");
        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeRectangle");
                logFileRec.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRec tuberec = new TubeRec();
                    tuberec.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberec.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberec.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberec.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tuberec.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    tuberec.Hauteur = Convert.ToDouble(getSqlNumericValue("HAUTEUR"));
                    tuberec.Largeur = Convert.ToDouble(getSqlNumericValue("LARGEUR"));
                    tuberec.epaisseur = Convert.ToDouble(getSqlNumericValue("EPAISSEUR"));
                    tuberec.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tuberec.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberec.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Recs.Add(tuberec);
                    logFileRec.Write("Tube : " + tuberec.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileRec.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        public override void WriteTubes()
        {
            if (List_Recs.Any())
            {
                foreach (TubeRec tuberec in List_Recs)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    string key = Guid.NewGuid().ToString();
                    //creation de la section
                    //sectionEntity = contextlocal.EntityManager.CreateEntity("_SECTION_CIRCLE");
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);
                    IEntity sectionQuality = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    string name = String.Format("Rec*{0}*{1}*{2}", tuberec.Largeur, tuberec.Hauteur, tuberec.epaisseur);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);// +"x"+ep.ToString());
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                    sectionEntity.SetFieldValue("_P_B", tuberec.Largeur);
                    sectionEntity.SetFieldValue("_P_H", tuberec.Hauteur);
                    sectionEntity.SetFieldValue("_P_T", tuberec.epaisseur);
                    //descpription //
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", String.Format("Rec hauteur={1} mm Larg={0} mm  Ep={2} mm", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur));
                    sectionEntity.Complete = true;
                    sectionEntity.Save();
                    logFileRec.Write("Tube : " + tuberec.COARTI + " sauvegarde");
                    //creation de la qualité
                    tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);
                    //recuperation de la section
                    if (SectionExclusion.Contains(name) == false)

                    {   //si n existe pas on recupere la nouvelle section //
                        sectionEntity.Save();
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                    }
                    else
                    //si elle existe deja on recupere la section //
                    { sectionQuality.SetFieldValue("_SECTION", getExistingSection(name, Section_Key)); }
                    if (TubeExclusion.Contains(tuberec.COARTI) == false)
                    {
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                        sectionQuality.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        sectionQuality.SetFieldValue("_BUY_COST", tuberec.PRIXART);
                        sectionQuality.Save();
                        logFileRec.Write("section " + name + " sauvegardée");
                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tuberec.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tuberec.Longueur);
                        barreEntity.SetFieldValue("_COMMENTS", string.Format("REC*{0}*{1}*{2}*{3}", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur, tuberec.Longueur));
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntity.GetImplementEntity("_SECTION").Id); //sectionEntity.GetImplementEntity("_SECTION").Id
                        barreEntity.Save();
                        logFileRec.Write("section " + name + "lng " + tuberec.Hauteur + " sauvegardée");
                    }
                }
            }

        }







    }


    /// <summary>
    /// recuperation des carres dans une liste de d'objet carres
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>



    public class Clipper_Import_Tube_Carre : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Rectangle;
        public double Diametre;
        public double Longueur;
        private List<TubeRec> List_Recs;

        TextWriterTraceListener logFileCar = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_REC_" + Properties.Resources.ImportTubeLog);



        public Clipper_Import_Tube_Carre(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_SHARP_RECTANGLE";
            SectionExclusion = getSectionExclusionList(Section_Key);
            TubeExclusion = getExclusionList("_BARTUBE", "_REFERENCE");
        }

        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.TubeCarre");
                logFileCar.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRec tuberec = new TubeRec();
                    tuberec.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberec.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberec.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberec.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tuberec.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    tuberec.Hauteur = Convert.ToDouble(getSqlNumericValue("LARGEUR"));
                    tuberec.Largeur = Convert.ToDouble(getSqlNumericValue("LARGEUR"));
                    tuberec.epaisseur = Convert.ToDouble(getSqlNumericValue("EPAISSEUR"));
                    tuberec.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tuberec.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberec.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Recs.Add(tuberec);
                    logFileCar.Write("Tube : " + tuberec.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileCar.Close();

                /*return listeTubeRond;*/
            }

        }


        //public void create(IContext contextlocal,string section_name, double diam, double ep, double lng, double cost)
        public override void WriteTubes()
        {
            if (List_Recs.Any())
            {
                foreach (TubeRec tuberec in List_Recs)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    string key = Guid.NewGuid().ToString();
                    //creation de la section
                    //sectionEntity = contextlocal.EntityManager.CreateEntity("_SECTION_CIRCLE");
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);
                    IEntity sectionQuality = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    string name = String.Format("Car*{0}*{2}", tuberec.Largeur, tuberec.Hauteur, tuberec.epaisseur);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);// +"x"+ep.ToString());
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                    sectionEntity.SetFieldValue("_P_B", tuberec.Largeur);
                    sectionEntity.SetFieldValue("_P_H", tuberec.Hauteur);
                    sectionEntity.SetFieldValue("_P_T", tuberec.epaisseur);
                    //descpription //
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", String.Format("Car hauteur={1} mm Larg={0} mm  Ep={2} mm", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur));
                    sectionEntity.Complete = true;
                    sectionEntity.Save();
                    logFileCar.Write("Tube : " + tuberec.COARTI + " sauvegarde");
                    //creation de la qualité
                    tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);

                    //recuperation de la section
                    if (SectionExclusion.Contains(name) == false)

                    {   //si n existe pas on recupere la nouvelle section //
                        sectionEntity.Save();
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                    }
                    else
                    //si elle existe deja on recupere la section //
                    { sectionQuality.SetFieldValue("_SECTION", getExistingSection(name, Section_Key)); }
                    if (TubeExclusion.Contains(tuberec.COARTI) == false)
                    {
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                        sectionQuality.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        sectionQuality.SetFieldValue("_BUY_COST", tuberec.PRIXART);

                        sectionQuality.Save();
                        logFileCar.Write("section " + name + " sauvegardée");
                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tuberec.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tuberec.Longueur);
                        barreEntity.SetFieldValue("_COMMENTS", string.Format("CAR*{0}*{2}*{3}", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur, tuberec.Longueur));
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntity.GetImplementEntity("_SECTION").Id); //sectionEntity.GetImplementEntity("_SECTION").Id
                        barreEntity.Save();
                        logFileCar.Write("section " + name + "lng " + tuberec.Hauteur + " sauvegardée");
                    }
                }
            }

        }















    }


    /// <summary>
    /// recuperation des flats dans une liste de d'objet flat
    /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
    /// les methodes virtuelles read : lise le contenu de lcipper avec le lien odbc 
    /// les methodes virtuelles write : ecrive les tubes selon la typologie
    /// </summary>


    public class Clipper_Import_Tube_Flat : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Flat;
        public double Diametre;
        public double Longueur;
        private List<TubeRec> List_Recs;

        public Clipper_Import_Tube_Flat(IContext context)
        {
            this.contextlocal = context;
            Section_Key = "_SECTION_FLAT";
            SectionExclusion = getSectionExclusionList(Section_Key);
            TubeExclusion = getExclusionList("_BARTUBE", "_REFERENCE");
        }

        TextWriterTraceListener logFileFlat = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_FLAT_" + Properties.Resources.ImportTubeLog);
        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// la liste TubeExclusion recupere les tubes deja importés et empeche la creation de doublons
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Recs = new List<TubeRec>();
                //recuperation des tube rectangluaires
                string sql_tube_rec = this.JSTOOLS.getJsonStringParametres("sql.Flat");
                logFileFlat.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_rec);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_rec;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeRec tuberec = new TubeRec();
                    tuberec.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tuberec.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tuberec.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tuberec.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tuberec.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    tuberec.Hauteur = Convert.ToDouble(getSqlNumericValue("HAUTEUR"));
                    tuberec.Largeur = Convert.ToDouble(getSqlNumericValue("LARGEUR"));
                    tuberec.epaisseur = 0;
                    tuberec.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tuberec.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tuberec.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Recs.Add(tuberec);
                    logFileFlat.Write("Tube : " + tuberec.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                List_Recs = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFileFlat.Close();

                /*return listeTubeRond;*/
            }

        }

        public override void WriteTubes()
        {
            if (List_Recs.Any())
            {
                foreach (TubeRec tuberec in List_Recs)
                {

                    //creation de la section (type...)//
                    IEntity sectionEntity;
                    string key = Guid.NewGuid().ToString();
                    //creation de la section
                    //sectionEntity = contextlocal.EntityManager.CreateEntity("_SECTION_CIRCLE");
                    sectionEntity = contextlocal.EntityManager.CreateEntity(Section_Key);
                    IEntity sectionQuality = contextlocal.EntityManager.CreateEntity("_SECTION_QUALITY");
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_KEY", key);
                    string name = String.Format("Flat*{0}*{2}", tuberec.Largeur, tuberec.Largeur, tuberec.Hauteur);
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_NAME", name);// +"x"+ep.ToString());
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_STANDARD", true);
                    sectionEntity.SetFieldValue("_P_B", tuberec.Largeur);
                    sectionEntity.SetFieldValue("_P_H", tuberec.Hauteur);
                    //sectionEntity.SetFieldValue("_P_T", tuberec.epaisseur);
                    //descpription //
                    sectionEntity.GetImplementEntity("_SECTION").SetFieldValue("_DESCRIPTION", String.Format("Flat  Largeur={1} mm hauteur={0} mm ", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur));
                    sectionEntity.Complete = true;
                    sectionEntity.Save();
                    logFileFlat.Write("Tube : " + tuberec.COARTI + " sauvegarde");
                    //creation de la qualité
                    tuberec.Material = tuberec.GetGrade(contextlocal, tuberec.Nuance, tuberec.Etat);

                    //recuperation de la section
                    if (SectionExclusion.Contains(name) == false)
                    {   //si n existe pas on recupere la nouvelle section //
                        sectionEntity.Save();
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                    }
                    else
                    //si elle existe deja on recupere la section //
                    { sectionQuality.SetFieldValue("_SECTION", getExistingSection(name, Section_Key)); }

                    if (TubeExclusion.Contains(tuberec.COARTI) == false)
                    {
                        sectionQuality.SetFieldValue("_SECTION", sectionEntity.Id32);
                        sectionQuality.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        sectionQuality.SetFieldValue("_BUY_COST", tuberec.PRIXART);
                        sectionQuality.Save();
                        logFileFlat.Write("section " + name + " sauvegardée");
                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tuberec.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tuberec.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tuberec.Longueur);
                        barreEntity.SetFieldValue("_COMMENTS", string.Format("FLAT*{0}*{1}*{3}", tuberec.Hauteur, tuberec.Largeur, tuberec.epaisseur, tuberec.Longueur));
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntity.GetImplementEntity("_SECTION").Id); //sectionEntity.GetImplementEntity("_SECTION").Id
                        barreEntity.Save();
                        logFileFlat.Write("section " + name + "lng " + tuberec.Hauteur + " sauvegardée");
                    }
                }
            }

        }















    }

    /// <summary>
    /// importe les tubes spéciaux : la racine du nom du tube doit contenir le nom de la section du tube spécial
    /// par exemple IPN80*S235*BRUT*3000 : la strucuture doit etre au minimum comme ceci: [SECTION ALMACAM].....[LONGUEUR] 
    /// ->[IPN80].....[3000] la section choisie est IPN80, la longeur est 3000
    /// </summary>

    public class Clipper_Import_Tube_Speciaux : Clipper_Article_Tube, IDisposable
    {
        //private List<TubeRond>listeTubeRond;
        private TypeTube type = TypeTube.Speciaux;
        private List<TubeSpe> List_Spe;


        public Clipper_Import_Tube_Speciaux(IContext context)
        {
            this.contextlocal = context;
            //string section_key;
            TubeExclusion = new List<string>();
            // Section_Key = typeTube; // "_SECTION_FLAT";
            //construction de la liste d'exclusion sur les section ci dessous
            getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE");

            /*
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_IPN"); 
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_IPE"); 
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_UPN"); 
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_UPE"); 
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_L"); 
              getExclusionList(ref TubeExclusion, "_BARTUBE", "_REFERENCE", "SECTION_LROUND"); */


        }

        /// <summary>
        /// utilise pour les tubes spes uniquement ou le section_key  (SECTION_UPE,SECTION_UPN..) peut varier
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="uid_field">chzamps unique de recherche</param>
        /// <param name="exclusionlist">list cumule d'exclusion.. </param>
        /// 
        /// <returns></returns>
        public void getExclusionList(ref List<string> exclusion, string entitype, string uid_field)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                //List<string> exclusion = new List<string>();
                //exclusionlist.ToList();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                if (exclusionlist.Count > 0)
                {
                    foreach (IEntity ex in exclusionlist)
                    {
                        if (ex.GetFieldValueAsString(uid_field) != null)
                        {
                            string uid = ex.GetFieldValueAsString(uid_field).Trim();
                            exclusion.Add(uid);

                        }




                    }

                }
                //return exclusion;
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                // return null;
            }


        }
        /// <summary>
        /// utilise pour les tubes spes uniquement ou le section_key  (SECTION_UPE,SECTION_UPN..) peut varier
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="uid_field">chzamps unique de recherche</param>
        /// <param name="section_key">Entité section SECTION_UPE,SECTION_UPN.. </param>
        /// <param name="exclusionlist">list cumule d'exclusion.. </param>
        /// 
        /// <returns></returns>
        public void getRestrictedExclusionList(ref List<string> exclusion, string entitype, string uid_field, string section_key)
        {
            try
            {

                IEntityList exclusionlist;
                exclusionlist = contextlocal.EntityManager.GetEntityList(entitype);
                //exclusionlist[0].GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key
                exclusionlist.Fill(false);
                //List<string> exclusion = new List<string>();
                //exclusionlist.ToList();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();

                if (exclusionlist.Count > 0)
                {
                    foreach (IEntity ex in exclusionlist)
                    {
                        if (ex.GetFieldValueAsString(uid_field) != null & ex.GetFieldValueAsEntity("_SECTION").ImplementedEntityType.Key == section_key)
                        {
                            string uid = ex.GetFieldValueAsString(uid_field).Trim();
                            exclusion.Add(uid);

                        }




                    }

                }
                //return exclusion;
            }

            catch (Exception ie)
            {

                MessageBox.Show(ie.Message);
                // return null;
            }


        }
        //traduit du code article les inforamtions en decomposant le code article
        private TubeSpe getTubeInfos(string coda)
        {

            try
            {

                TubeSpe sp = new TubeSpe();
                string[] infos = null;

                if (coda != null || coda != string.Empty)
                    //recuperation de la section spe du tube

                    infos = coda.Split('*');
                long dim = infos.Count() - 1;
                sp.Section = infos[0].Trim();
                sp.Longueur = Convert.ToDouble(infos[dim]);
                //section
                if (sp.Section.Contains("IPN")) { sp.Section_Spe_Key = "_SECTION_IPN"; }
                else if (sp.Section.Contains("IPE")) { sp.Section_Spe_Key = "_SECTION_IPE"; }
                else if (sp.Section.Contains("UPN")) { sp.Section_Spe_Key = "_SECTION_UPN"; }
                else if (sp.Section.Contains("UPE")) { sp.Section_Spe_Key = "_SECTION_UPE"; }
                else if (sp.Section == "L") { sp.Section_Spe_Key = "_SECTION_L"; }
                else if (sp.Section.Contains("LROUND")) { sp.Section_Spe_Key = "_SECTION_LROUND"; }
                else { sp.Section_Spe_Key = string.Empty; }
                sp.COARTI = coda;
                return sp;

            }

            catch (Exception ie)
            {
                MessageBox.Show(ie.Message); return null;
            
                
            }

        }

        TextWriterTraceListener logFilSpe = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_SPE_" + Properties.Resources.ImportTubeLog);
        /// <summary>
        /// recuperation des ronds dans une liste de d'objet rond
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void ReadTubes()
        {
            //creation de la liste des tube ronds
            //List<TubeRond> listeTubeRond;//= new List<TubeRond>();

            try
            {
                //creation de la liste des tube ronds
                //listeTubeRond = new List<TubeRond>();
                this.List_Spe = new List<TubeSpe>();
                //recuperation des tube rectangluaires
                string sql_tube_spe = this.JSTOOLS.getJsonStringParametres("sql.Profilspeciaux");
                logFilSpe.Write("requete utilisée pour l'import des tubes \r\n " + sql_tube_spe);
                int ii = 0;
                this.DbCommand.CommandText = sql_tube_spe;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    TubeSpe tubespe = new TubeSpe();
                    //tubespe.COARTI = TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    tubespe = getTubeInfos(TABLE_ARTICLEM["COARTI"].ToString().Trim());
                    tubespe.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    tubespe.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    tubespe.PRIXART = Convert.ToDouble(getSqlNumericValue("PRIXART"));
                    //tuberond.Thickness = Convert.ToDouble(TABLE_ARTICLEM["EPAISSEUR"]);
                    tubespe.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));

                    tubespe.Longueur = Convert.ToDouble(getSqlNumericValue("LNG"));
                    tubespe.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    tubespe.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    this.List_Spe.Add(tubespe);
                    logFilSpe.Write("Tube : " + tubespe.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                //logFileRond.Flush();
                //logFileRond.Close();
                //return listeTubeRond;
            }

            catch (Exception ie)
            {

                List_Spe = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFilSpe.Close();

                /*return listeTubeRond;*/
            }

        }



        public override void WriteTubes()
        {
            if (List_Spe.Any())
            {
                //liste des type actuelles



                foreach (TubeSpe tubespe in List_Spe)
                {

                    //creation de la section (type...)//
                    /*
                    IEntityList sectionEntityList;
                    IEntity sectionEntity;
                    sectionEntityList = contextlocal.EntityManager.GetEntityList(tubespe.COARTI);
                    sectionEntityList.Fill(false);
                    sectionEntity = sectionEntityList.FirstOrDefault();*/

                    //+ @"\IMPLEMENTED_SECTION", "_NAME", ConditionOperator.Equal, tubespe.Section
                    //logFilSpe.Write("Tube : " + tubespe.COARTI + " trouvée");
                    //onne creer que ce qui n'existe pas


                    if (TubeExclusion.Contains(tubespe.COARTI) == false)
                    {
                        //si elle existe deja on recupere la section //
                        //IEntityList sectionQualityList, sectionEntityList;
                        //IEntity sectionQuality, sectionEntity;
                        long sectionEntityid;

                        //creation de la qualité
                        tubespe.Material = tubespe.GetGrade(contextlocal, tubespe.Nuance, tubespe.Etat);
                        //sectionQualityList = contextlocal.EntityManager.GetEntityList(tubespe.Material.GetFieldValueAsString("_NAME"));
                        //sectionQualityList.Fill(false);
                        //sectionQuality = sectionQualityList.FirstOrDefault();
                        //recuperation de la section
                        sectionEntityid = getExistingImplementedSection(tubespe.Section, tubespe.Section_Spe_Key);



                        //creation DE LA BARRE
                        IEntity barreEntity;
                        string keybar = Guid.NewGuid().ToString();
                        barreEntity = contextlocal.EntityManager.CreateEntity("_BARTUBE");
                        barreEntity.SetFieldValue("_REFERENCE", tubespe.COARTI);
                        barreEntity.SetFieldValue("_QUALITY", tubespe.Material.Id32);
                        barreEntity.SetFieldValue("_LENGTH", tubespe.Longueur);
                        barreEntity.SetFieldValue("_COMMENTS", string.Format("section: {0} matiere: {1} longeur: {2}", tubespe.Section, tubespe.Material.GetFieldValueAsString("_NAME"), tubespe.Longueur));
                        //ATTENTION recuperation des infos de l'implemented section
                        barreEntity.SetFieldValue("_SECTION", sectionEntityid); //sectionEntity.GetImplementEntity("_SECTION").Id

                        barreEntity.Save();
                        logFilSpe.Write("section " + tubespe.Section + "lng " + tubespe.Longueur + " sauvegardée");
                    }


                }
            }
        }


    }


    #endregion


    #region Founitures
   

    /// <summary>
    /// import specifique des vis
    /// </summary>
    public class Clipper_Import_Fournitures_Divers : Clipper_Article, IDisposable
    {
        //
        // public double Diametre;
        //public double Longueur;

        private List<Founiture_Divers> List_Fourniture;

        public string Key; //"_SECTION_CIRCLE" or....
        //public List<string> SectionExclusion = new List<string>();
        public Dictionary<long, string> Founiture_Divers_Exclusion = new Dictionary<long, string>();

        TextWriterTraceListener logFourniture = new TextWriterTraceListener(System.IO.Path.GetTempPath() + "\\" + "_SIMPLE_SUPPLY_" + Properties.Resources.ImportFournitureLog);

        /// <summary>
        /// constructeur obligatoire pour ajuster les elements a controler
        /// </summary>
        /// <param name="context"></param>
        public Clipper_Import_Fournitures_Divers(IContext context)
        {

            

            this.contextlocal = context;
            contextlocal.TraceLogger.TraceInformation("Test de Connection ODBC");
            Key = "_SIMPLE_SUPPLY";
            Founiture_Divers_Exclusion = getExclusionList(Key);

            Odbc_Connexion();
            contextlocal.TraceLogger.TraceInformation("Connection ODBC ok");
        }



        /// <summary>
        /// retourn une liste d'exclusion des sections ou des tubes
        /// </summary>
        /// <param name="entitype"></param>
        /// <param name="entity_uniquestring_field"></param>
        /// <returns></returns>
        public Dictionary<long, string> getExclusionList(string entitypekey)
        {

            try
            {
                contextlocal.TraceLogger.TraceInformation("creation de la liste d'exclusion");
                IEntityList exclusionlist;

                exclusionlist = contextlocal.EntityManager.GetEntityList(entitypekey);
                exclusionlist.Fill(false);
                Dictionary<long,string> exclusion = new Dictionary<long, string>();
                // Find material           
                //Entity currentQuality = qualitylist.Where(q => q.DefaultValue.Equals(quality)).FirstOrDefault();
                foreach (IEntity ex in exclusionlist)
                {

                    string coarti = ex.GetImplementEntity("_SUPPLY").GetFieldValueAsString("_REFERENCE").Trim();
                    long uid = ex.GetImplementEntity("_SUPPLY").Id;
                    //ex.GetImplementEntity(Key);

                    if (exclusion.ContainsKey(uid) == false && string.IsNullOrEmpty(coarti) == false)
                    {
                        exclusion.Add(uid,coarti);
                    }

                }

                contextlocal.TraceLogger.TraceInformation("creation de la liste d'exclusion terminée");
                return exclusion;
            }

            catch
            {

                return null;
            }
        }


        /// <summary>
        /// recuperation des ronds dans une liste de d'objet vis de la base clipper
        /// </summary>
        /// <returns>List<TubeRond></returns>
        public override void Read()
        {
            contextlocal.TraceLogger.TraceInformation("lecture des fournitures clipper en cours");

            try
            {
                //creation de la liste des fournitures
          
                this.List_Fourniture = new List<Founiture_Divers>();
                //recuperation des tube rectangluaires
                string sql_vis = this.JSTOOLS.getJsonStringParametres("sql.Divers");
                logFourniture.Write("requete utilisée pour l'import des fournitures \r\n " + sql_vis);
                int ii = 0;
                this.DbCommand.CommandText = sql_vis;

                TABLE_ARTICLEM = DbCommand.ExecuteReader();

                while (TABLE_ARTICLEM.Read())
                {
                    ii++;
                    Founiture_Divers Fourniture = new Founiture_Divers();
                    ///cle unique clipper
                    Fourniture.AMCLEUNIK= Convert.ToInt64( TABLE_ARTICLEM["AMCLEUNIK"]);
                    Fourniture.COARTI= TABLE_ARTICLEM["COARTI"].ToString().Trim();
                    Fourniture.Nuance = TABLE_ARTICLEM["CODENUANCE"].ToString().Trim(); //CODEETAT, Tech_NuanceMatiere.Nuance AS CODENUANCE
                    Fourniture.Etat = TABLE_ARTICLEM["CODEETAT"].ToString().Trim();
                    Fourniture.Densite = Convert.ToDouble(getSqlNumericValue("DENSITE"));
                    Fourniture.UnitePrix = TABLE_ARTICLEM["UnitePrix"].ToString().Trim();
                    Fourniture.PRIXART = Math.Round(GetPrice(Fourniture.UnitePrix, Convert.ToDouble(getSqlNumericValue("PRIXART"))), 5);
                    Fourniture.COFA = TABLE_ARTICLEM["COFA"].ToString().Trim(); ;
                    Fourniture.Name = TABLE_ARTICLEM["COARTI"].ToString().Trim() + "*" + getSqlNumericValue("LNG");
                    Fourniture.DESA1 = TABLE_ARTICLEM["DESA1"].ToString().Trim();
                    
                    this.List_Fourniture.Add(Fourniture);
                    logFourniture.Write("Fourniture : " + Fourniture.COARTI + " capturé");

                };

                TABLE_ARTICLEM.Close();
                contextlocal.TraceLogger.TraceInformation("lecture des fournitures terminée");
                //return listeTubeRond;
            }

            catch (Exception ie)
            {
                this.List_Fourniture = null;
                MessageBox.Show(ie.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                logFourniture.Close();


            }

        }

        /// <summary>
        /// ecriture dans la base almacam
        /// </summary>
        /// 
        //ecriture des fournitures

        public override void Write()
        {
            try
            {
                contextlocal.TraceLogger.TraceInformation("Import des fournitures en cours");

                if (this.List_Fourniture.Any())
                {
                  

                    foreach (Founiture_Divers fourniture in this.List_Fourniture)
                    {   

                        if (!Founiture_Divers_Exclusion.ContainsValue( fourniture.COARTI) ) {
                         IEntity Fourniture_Entity;
//
                        Fourniture_Entity = contextlocal.EntityManager.CreateEntity(Key);
                        Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_REFERENCE", fourniture.COARTI);
                        Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_DESIGNATION", fourniture.DESA1 );
                        Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_COMMENTS", fourniture.COFA);
                        Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_BUY_COST", fourniture.PRIXART);
                        Fourniture_Entity.Save();
                        }
                        //update
                        else {
                            IEntity Fourniture_Entity;
                            //recupere l'id et mets a jour les champs
                            //FirstOrDefault(x => x.Value == "one").Key; 
                            long uid =Founiture_Divers_Exclusion.FirstOrDefault(x => x.Value == fourniture.COARTI).Key;
                            Fourniture_Entity = contextlocal.EntityManager.GetEntity(uid, "_SUPPLY");
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_REFERENCE", fourniture.COARTI);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_DESIGNATION", fourniture.DESA1);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_COMMENTS", fourniture.COFA);
                            Fourniture_Entity.GetImplementEntity("_SUPPLY").SetFieldValue("_BUY_COST", fourniture.PRIXART);

                        }
                    }


                }



                contextlocal.TraceLogger.TraceInformation("Import des fornitures terminé avec succes");

            }


            catch
            {

            }

        }


        /// <summary>
        /// ecriture dans la base almacam
        /// </summary>
        /// 
        //ecriturze des vis

        public override void Update()
        {
            try
            {
                if (this.List_Fourniture.Any())
                {

                    foreach (Founiture_Divers fourniture in this.List_Fourniture)
                    {
                        IEntity Fourniture_Entity;

                     
                        Fourniture_Entity = contextlocal.EntityManager.CreateEntity("_SUPPLY");
                        Fourniture_Entity.SetFieldValue("_REFERENCE", fourniture.Name);
                        Fourniture_Entity.SetFieldValue("_COMMENTS", "");
                      
                        Fourniture_Entity.Save();
                    }


                }

            }


            catch
            {

            }

        }




    }

    #endregion



}

#endregion

#region Exception
public class MissingJsonFile : Exception
{

    public void MissingJsonFileException(string parametername)
    {
        {

            MessageBox.Show("Il manque le fichier Json " + parametername + " dans le repertoire almacam");


        }

    }

}

[Serializable]
internal class Missing_Obdc_Exception : Exception

{
   
      
    
        public Missing_Obdc_Exception( string nomodbc)

        {

            MessageBox.Show("la connexion odbc "+ nomodbc + " n'est pas configurée. Contactez Clipper pour ce point ou bien referez vous a la documentation d'installaitons");
            Environment.Exit(0);

        }

       

    
}



#endregion

#region Tools
public static class AlmaCamTool
{
    
    /// <summary>
    /// recupere un context almacam
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static IContext GetContext(IContext context)
    {
        try
        {

            if (context == null)
            {



                string DbName = Alma_RegitryInfos.GetLastDataBase();
                IModelsRepository modelsRepository = new ModelsRepository();
                IContext contextlocal = modelsRepository.GetModelContext(DbName);
                return contextlocal;

            }
            else
            {
                return context;
            }
        }
        catch (Exception ie) {

            string methode = System.Reflection.MethodBase.GetCurrentMethod().Name;
            MessageBox.Show(ie.Message);

            return null; }
    }

    /// <summary>
   /// verifie si la connexion odbc est possible
   /// </summary>
   /// <param name="DSN"></param>
   /// <returns></returns>
    public static bool Is_Odbc_Exists(string DSN)
    {
        bool rst = true;
        var mDbConnection = new OdbcConnection("DSN=" + DSN);

        try {
            

            
         
            using (mDbConnection)
            {
                mDbConnection.Open();

                rst = true;
            }
            

            return rst;


        } catch(Exception ie)
        {
            return false;
        }
        finally { mDbConnection.Close(); }
    }
}
#endregion