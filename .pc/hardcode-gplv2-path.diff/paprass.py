#!/usr/bin/env python
#-*- coding: utf-8 -*-

import wx
import wx.calendar
import Image
import ImageEnhance
from datetime import date
import ConfigParser
import os
import sys, shutil
version = sys.version_info[0] + sys.version_info[1] * 0.1
if version < 2.5:
    from pysqlite2 import dbapi2 as sqlite
else:
    import sqlite3 as sqlite
if "win" in sys.platform:
    WIN = True
else:
    WIN = False
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
if not WIN:
    import sane
from xml.dom.minidom import Document

# Pap'rass - copyright 2007-2008 Alain Delgrange aka bipede
# GED personnelle développée sous licence GPL V2

def IsItImage(racine):
    for term in GLOBVAR.listedoc:
        if term != "PDF":
            if os.path.isfile(racine + "." + term.lower()):
                return racine + "." + term.lower()
    return False
    
def IsItOoffice(racine):
    for term in GLOBVAR.listeoo:
        if os.path.isfile(racine + "." + term.lower()):
            return racine + "." + term.lower()
    return False

def upper(chaine):
    return chaine.upper()

class GlobalVar:
    def __init__(self):
        self.appdir = os.path.split(os.path.abspath(sys.argv[0]))[0]
        self.homedir = os.path.join(os.path.expanduser("~"), ".paprass-data")
        self.datadir = os.path.join(self.homedir, "data")
        self.docdir = os.path.join(self.datadir, "documents")
        self.tempdir = os.path.join(self.datadir, "temp")
        self.listedoc = ["JPG", "PNG", "PDF", "TIF", "BMP", "PNM"]
        self.listeoo = ["ODT", "ODS", "ODP", "ODG"]
        self.themedir = "INDEFINI"
        self.base = "INDEFINI"
        self.visupdf = "INDEFINI"
        self.app = None

class ScanVar:
    def __init__(self):
        self.device = "INDEFINI"
        self.mode = "INDEFINI"
        self.resolution = "INDEFINI"

    def SetDevice(self, device):
        self.device = device

    def SetMode(self, mode):
        self.mode = mode

    def SetResolution(self, resolution):
        self.resolution = resolution

class Config:
    def __init__(self):
        if not os.path.isdir(GLOBVAR.homedir):
            os.mkdir(GLOBVAR.homedir)
        self.file_cfg = os.path.join(GLOBVAR.homedir, ".paprass.cfg")
        self.connec = None
        config = ConfigParser.ConfigParser()
        if len(config.read([self.file_cfg])) == 0:
            config.add_section("viewer")
            if WIN:
                config.set("viewer", "pdf", "")
            else:
                config.set("viewer", "pdf", "evince")
            config.add_section("looknfeel")
            config.set("looknfeel", "theme", "tangerine")
            config.add_section("scanner")
            config.set("scanner", "device", "INDEFINI")
            config.set("scanner", "mode", "INDEFINI")
            config.set("scanner", "resolution", "INDEFINI")
            file_cfg = open(self.file_cfg, 'wb')
            config.write(file_cfg)
            file_cfg.close()
        self.viewer = config.get("viewer", "pdf")
        self.theme = config.get("looknfeel", "theme")
        try:
            self.device = config.get("scanner", "device")
        except:
            config.add_section("scanner")
            config.set("scanner", "device", "INDEFINI")
            config.set("scanner", "mode", "INDEFINI")
            config.set("scanner", "resolution", "INDEFINI")
            file_cfg = open(self.file_cfg, 'wb')
            config.write(file_cfg)
            file_cfg.close()
            self.device = config.get("scanner", "device")
        self.mode = config.get("scanner", "mode")
        self.resolution = config.get("scanner", "resolution")
        if not os.path.isdir(GLOBVAR.datadir):
            os.mkdir(GLOBVAR.datadir)
            os.mkdir(GLOBVAR.docdir)
            os.mkdir(GLOBVAR.tempdir)
            base = os.path.join(GLOBVAR.datadir, "data.ged")
            self.connec = sqlite.connect(base, isolation_level=None)
            cursor = self.connec.cursor()
            req = "CREATE TABLE classeurs (classeur INTEGER PRIMARY KEY AUTOINCREMENT, libelle TEXT)"
            cursor.execute(req)
            req = "CREATE TABLE dossiers (dossier INTEGER PRIMARY KEY AUTOINCREMENT, classeur INTEGER, libelle TEXT)"
            cursor.execute(req)
            req = "CREATE TABLE chemises (chemise INTEGER PRIMARY KEY AUTOINCREMENT, classeur INTEGER, dossier INTEGER, libelle TEXT)"
            cursor.execute(req)
            req = "CREATE TABLE documents (enreg INTEGER PRIMARY KEY AUTOINCREMENT, "
            req = req + "classeur INTEGER, dossier INTEGER, chemise INTEGER, date TEXT, titre TEXT, nbpages INTEGER, annee TEXT, mois TEXT)"
            cursor.execute(req)
        else:
            base = os.path.join(GLOBVAR.datadir, "data.ged")
            self.connec = sqlite.connect(base, isolation_level=None)
            self.connec.create_function("majuscules", 1, upper)
            c = self.connec.cursor()
            c.execute("VACUUM")

    def GetBase(self):
        return self.connec

    def GetViewer(self):
        return self.viewer

    def GetTheme(self):
        return self.theme

    def GetDevice(self):
        return self.device

    def GetMode(self):
        return self.mode

    def GetResolution(self):
        return self.resolution

    def SetTheme(self, theme):
        config = ConfigParser.ConfigParser()
        config.read([self.file_cfg])
        config.set("looknfeel", "theme", theme)
        file_cfg = open(self.file_cfg, 'wb')
        config.write(file_cfg)
        file_cfg.close()
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               message=u"Le thème %s sera pris en compte au prochain lancement de Pap'rass"%theme,
                               caption=u"Changement de thème",
                               style=wx.OK|wx.ICON_INFORMATION)
        res = dlg.ShowModal()
        dlg.Destroy()

    def SetViewer(self, viewer):
        config = ConfigParser.ConfigParser()
        config.read([self.file_cfg])
        config.set("viewer", "pdf", viewer)
        file_cfg = open(self.file_cfg, 'wb')
        config.write(file_cfg)
        file_cfg.close()
        GLOBVAR.visupdf = viewer
        dlg = MessageDialog(parent=GLOBVAR.app,
                            caption=u"Changement de viewer PDF",
                            message=u"Le viewer %s est maintenant pris en compte par Pap'rass"%viewer,
                            style=wx.OK|wxICON_INFORMATION)
        res = dlg.ShowModal()
        dlg.Destroy()

    def SetScanner(self, device):
        config = ConfigParser.ConfigParser()
        config.read([self.file_cfg])
        config.set("scanner", "device", device)
        file_cfg = open(self.file_cfg, 'wb')
        config.write(file_cfg)
        file_cfg.close()

    def SetConfigScanner(self, mode, resolution):
        config = ConfigParser.ConfigParser()
        config.read([self.file_cfg])
        config.set("scanner", "mode", mode)
        config.set("scanner", "resolution", resolution)
        file_cfg = open(self.file_cfg, 'wb')
        config.write(file_cfg)
        file_cfg.close()

class ConfigParam(wx.Dialog):
    def __init__(self, parent = None):
        wx.Dialog.__init__(self, parent = parent, id = -1, title = u"Modification des paramètres")
        box0 = wx.BoxSizer(wx.VERTICAL)
        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        box3 = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(self, -1, u"Choisir le thème graphique :  ")
        self.theme_list = ["classic", "tangerine", "chocolate", "blue", "green"]
        self.combo = wx.ComboBox(self, -1, size= (100, -1),
                                value=CONFIG.GetTheme(),
                                choices = self.theme_list,
                                style = wx.CB_READONLY)
        box2.Add(label, 4, flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        box2.Add((0,0),1)
        box2.Add(self.combo, 4, flag=wx.ALIGN_RIGHT)
        label = wx.StaticText(self, -1, u"Saisir le nom du viewer PDF : ")
        self.viewer_entry = wx.TextCtrl(self, -1, size= (100, -1))
        if WIN:
            self.viewer_entry.SetValue("")
            self.viewer_entry.Enable(False)
        else:
            self.viewer_entry.SetValue(CONFIG.GetViewer())
        box3.Add(label, 4, flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL)
        box3.Add((0,0),1)
        box3.Add(self.viewer_entry, 4, flag=wx.ALIGN_RIGHT)
        box1.Add(box2, flag=wx.ALL, border = 5)
        box1.Add(box3, flag=wx.ALL, border = 5)
        sizerbouton = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box1.Add(sizerbouton, flag=wx.ALIGN_CENTER|wx.TOP, border=10)
        box0.Add(box1, flag = wx.ALL, border=20)
        self.SetSizer(box0)
        self.Fit()
        self.CentreOnParent()

    def GetTheme(self):
        return self.theme_list[self.combo.GetSelection()]

    def GetViewer(self):
        return self.viewer_entry.GetValue()

class Configuration(wx.Panel):
    def __init__(self, parent):
        self.conteneur = parent
        larg, haut = parent.GetClientSizeTuple()
        taille = wx.Size(larg, haut)
        wx.Panel.__init__(self, parent=parent, id=-1, size=taille)
        self.plan = False
        self.ecran = None
        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        self.c = GLOBVAR.base.cursor()
        larg_boutons = (larg-20)/3
        self.btClassement = wx.Button(self, -1, u'Définir le plan de classement')
        self.btClassement.Bind(wx.EVT_BUTTON, self.DefinirClassement)
        self.btThemes = wx.Button(self, -1, u'Changer de thème ou de viewer PDF')
        self.btThemes.Bind(wx.EVT_BUTTON, self.ChangerTheme)
        self.btExport = wx.Button(self, -1, u'Exporter la base des documents en XML')
        self.btExport.Bind(wx.EVT_BUTTON, self.ExporterBase)
        box2.Add((0,0),1)
        box2.Add(self.btClassement, 20)
        box2.Add((0,0),1)
        box2.Add(self.btThemes, 20)
        box2.Add((0,0),1)
        box2.Add(self.btExport, 20)
        box2.Add((0,0),1)
        self.vue = wx.Panel(self, -1,)
        box1.Add(box2, flag=wx.EXPAND)
        box1.Add(self.vue, 1, flag=wx.EXPAND)
        self.vueSizer = wx.BoxSizer(wx.VERTICAL)
        self.vue.SetSizer(self.vueSizer)
        self.vue.SetAutoLayout(True)
        self.SetSizer(box1)
        box1.Fit(self)

    def ExporterBase(self, event):
        entete = """<?xml version="1.0" encoding= "utf-8"?>
                    <!DOCTYPE documents [
                    <!ELEMENT documents (document+)>
                    <!ELEMENT document (classement?,titre,dateISO,pages)>
                    <!ELEMENT classement (classeur,dossier,chemise)>
                    <!ELEMENT classeur (#PCDATA)>
                    <!ELEMENT dossier (#PCDATA)>
                    <!ELEMENT chemise (#PCDATA)>
                    <!ELEMENT titre (#PCDATA)>
                    <!ELEMENT dateISO (#PCDATA)>
                    <!ELEMENT pages (page+)>
                    <!ELEMENT page (#PCDATA)>
                    <!ATTLIST page numero CDATA #REQUIRED>
                    ]>"""

        req = "SELECT enreg, titre, date, annee, mois, nbpages, classeur, dossier, chemise FROM documents ORDER BY enreg"
        res = self.c.execute(req)
        liste = res.fetchall()
        docu = Document()
        racine = docu.createElement("documents")
        docu.appendChild(racine)
        for e in liste:
            element = docu.createElement("document")
            racine.appendChild(element)
            titre = docu.createElement("titre")
            element.appendChild(titre)
            texte = docu.createTextNode(e[1])
            titre.appendChild(texte)
            dateISO = docu.createElement("dateISO")
            element.appendChild(dateISO)
            texte = docu.createTextNode(e[2])
            dateISO.appendChild(texte)
            if e[6] != 0:
                classement = docu.createElement("classement")
                element.appendChild(classement)
                req = "SELECT libelle FROM classeurs WHERE classeur = " + str(e[6])
                res = self.c.execute(req)
                libelle = res.fetchone()[0]
                classeur = docu.createElement("classeur")
                classement.appendChild(classeur)
                texte = docu.createTextNode(libelle)
                classeur.appendChild(texte)
                req = "SELECT libelle FROM dossiers WHERE classeur = " + str(e[6]) + " and dossier = " + str(e[7])
                res = self.c.execute(req)
                libelle = res.fetchone()[0]
                dossier = docu.createElement("dossier")
                classement.appendChild(dossier)
                texte = docu.createTextNode(libelle)
                dossier.appendChild(texte)
                req = "SELECT libelle FROM chemises WHERE classeur = " + str(e[6]) + " and dossier = " + str(e[7]) + " and chemise = " + str(e[8])
                res = self.c.execute(req)
                libelle = res.fetchone()[0]
                chemise = docu.createElement("chemise")
                classement.appendChild(chemise)
                texte = docu.createTextNode(libelle)
                chemise.appendChild(texte)
            pages = docu.createElement("pages")
            element.appendChild(pages)
            num = str(e[0])
            for page in range(e[5]):
                p = page + 1
                chemin = os.path.join(GLOBVAR.docdir, e[3], e[4])
                fic = num + "-" + str(p) + ".pdf"
                if os.path.isfile(os.path.join(chemin, fic)):
                    lapage = os.path.join(chemin, fic)
                else:
                    fic = num + "-" + str(p) + ".txt"
                    if os.path.isfile(os.path.join(chemin, fic)):
                        lapage = os.path.join(chemin, fic)
                    else:
                        fic = num + "-" + str(p)
                        if IsItOoffice(os.path.join(chemin, fic)):
                            lapage = IsItOoffice(os.path.join(chemin, fic))
                        else:
                            lapage = IsItImage(os.path.join(chemin, fic))
                noeud = docu.createElement("page")
                noeud.setAttribute("numero", str(p))
                pages.appendChild(noeud)
                texte = docu.createTextNode(lapage)
                noeud.appendChild(texte)
        fichier = os.path.join(GLOBVAR.homedir, "exportpaprass.xml")
        file = open(fichier, "wb")
        file.write(entete + (docu.toprettyxml().encode("utf-8").split('<?xml version="1.0" ?>')[1]))
        file.close()
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                            message=u"Le fichier exportpaprass.xml été constitué\net placé dans le répertoire principal de l'application",
                            caption=u"Exportation réalisée",
                            style=wx.OK|wx.ICON_INFORMATION)
        val = dlg.ShowModal()
        dlg.Destroy()

    def DefinirClassement(self, event):
        if not self.plan:
            self.plan = True
            self.btClassement.SetLabel(u"Terminer")
            self.btThemes.Enable(False)
            self.btExport.Enable(False)
            GLOBVAR.app.barre.EnableTool(GLOBVAR.app.ID_FERMER, False)
            GLOBVAR.app.menuDoc.Enable(GLOBVAR.app.ID_FERMER, False)
            self.ecran = PlanDeClassement(self.vue)
            self.vueSizer.Add(self.ecran, 1, wx.EXPAND)
            self.vueSizer.Fit(self.vue)
            self.Layout()
        else:
            self.plan = False
            self.ecran.Destroy()
            self.btClassement.SetLabel(u"Définir le plan de classement")
            self.btThemes.Enable(True)
            self.btExport.Enable(True)
            GLOBVAR.app.barre.EnableTool(GLOBVAR.app.ID_FERMER, True)
            GLOBVAR.app.menuDoc.Enable(GLOBVAR.app.ID_FERMER, True)
            self.Layout()

    def ChangerTheme(self, event):
        dlg = ConfigParam(GLOBVAR.app)
        res = dlg.ShowModal()
        theme = dlg.GetTheme()
        viewer = dlg.GetViewer()
        dlg.Destroy()
        if res == wx.ID_OK:
            if theme != CONFIG.GetTheme():
                CONFIG.SetTheme(theme)
            if viewer != "" and viewer != CONFIG.GetViewer():
                CONFIG.SetViewer(viewer)

class PanneauLogo(wx.Panel):
    def __init__(self, parent):
        taille = parent.GetClientSize()
        wx.Panel.__init__(self, parent, -1)
        self.SetBackgroundColour(wx.WHITE)
        sizer = wx.BoxSizer(wx.VERTICAL)
        image = wx.StaticBitmap(self, -1,
                                wx.Bitmap(os.path.join(GLOBVAR.themedir, "logo.png"),
                                wx.BITMAP_TYPE_PNG))
        sizer.Add((0,0),1)
        sizer.Add(image, 1, flag=wx.CENTRE)
        sizer.Add((0,0),1)
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        wx.EVT_SIZE(self, self.OnSize)

    def OnSize(self, event):
        self.Layout()

class PlanDeClassement(wx.Panel):
    def __init__(self, parent):
        (largeur, hauteur) = parent.GetClientSizeTuple()
        taille = wx.Size(largeur, hauteur)
        wx.Panel.__init__(self, parent=parent, id=-1, size=taille)
        self.valeur = []
        self.itemChoisi = None
        self.iter = None
        self.fenetre = wx.ScrolledWindow(self, -1)
        fenSizer = wx.BoxSizer(wx.VERTICAL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.listeImage = wx.ImageList(24, 24)
        self.imRoot = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "racine.png"), wx.BITMAP_TYPE_PNG))
        self.imClasseur = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "classeur.png"), wx.BITMAP_TYPE_PNG))
        self.imDossier = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "dossier.png"), wx.BITMAP_TYPE_PNG))
        self.imChemise = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "chemise.png"), wx.BITMAP_TYPE_PNG))
        self.texte = wx.StaticText(self,
                                   -1,
                                   label = u"Cliquez sur un item avec le bouton droit de la souris pour faire apparaître le menu déroulant correspondant",
                                   style = wx.ALIGN_CENTER | wx.ST_NO_AUTORESIZE)
        sizer2.Add((0,0), 0)
        sizer2.Add(self.texte, 1, wx.EXPAND)
        sizer2.Add((0, 0), 0)
        self.arbre = wx.TreeCtrl(self.fenetre, -1)
        fenSizer.Add(self.arbre, 1, wx.EXPAND)
        self.fenetre.SetSizer(fenSizer)
        self.fenetre.SetAutoLayout(True)
        self.arbre.SetImageList(self.listeImage)
        self.popMenuRoot = wx.Menu()
        self.popMenuClasseur = wx.Menu()
        self.popMenuDossier = wx.Menu()
        self.popMenuChemise = wx.Menu()

        self.popMenuRoot.Append(wx.ID_FILE1, u"Ajouter un classeur")

        self.popMenuClasseur.Append(wx.ID_FILE2, u"Ajouter un dossier au classeur")
        self.popMenuClasseur.Append(wx.ID_FILE7, u"Renommer le classeur")
        self.popMenuClasseur.Append(wx.ID_FILE3, u"Supprimer le classeur")

        self.popMenuDossier.Append(wx.ID_FILE4, u"Ajouter une chemise")
        self.popMenuDossier.Append(wx.ID_FILE8, u"Renommer le dossier")
        self.popMenuDossier.Append(wx.ID_FILE5, u"Supprimer le dossier")

        self.popMenuChemise.Append(wx.ID_FILE9, u"Renommer la chemise")
        self.popMenuChemise.Append(wx.ID_FILE6, u"Supprimer la chemise")
        self.c = GLOBVAR.base.cursor()
        self.Remplir()
        sizer.Add(sizer2, flag=wx.EXPAND|wx.ALL, border=5)
        sizer.Add(self.fenetre, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.SetAutoLayout(True)

        self.fenetre.SetScrollRate(20, 20)

        self.arbre.Bind(wx.EVT_TREE_ITEM_RIGHT_CLICK, self.ClickDroit)

        self.Bind(wx.EVT_MENU, self.AjoutClasseur, id = wx.ID_FILE1)

        self.Bind(wx.EVT_MENU, self.AjoutDossier, id = wx.ID_FILE2)
        self.Bind(wx.EVT_MENU, self.RenommeClasseur, id = wx.ID_FILE7)
        self.Bind(wx.EVT_MENU, self.SupprimeClasseur, id = wx.ID_FILE3)

        self.Bind(wx.EVT_MENU, self.AjoutChemise, id = wx.ID_FILE4)
        self.Bind(wx.EVT_MENU, self.RenommeDossier, id = wx.ID_FILE8)
        self.Bind(wx.EVT_MENU, self.SupprimeDossier, id = wx.ID_FILE5)

        self.Bind(wx.EVT_MENU, self.RenommeChemise, id = wx.ID_FILE9)
        self.Bind(wx.EVT_MENU, self.SupprimeChemise, id = wx.ID_FILE6)

    def Remplir(self):
        self.root = self.arbre.AddRoot(u"Plan de classement des documents", self.imRoot)
        self.itemChoisi = self.root

        self.c.execute("SELECT COUNT(*) FROM classeurs")
        if (self.c.fetchall()[0][0] > 0):
            self.c.execute("SELECT classeur, libelle FROM classeurs ORDER BY classeur")
            listeClasseurs = self.c.fetchall()

            for x in listeClasseurs:
                leClasseur = x[0]
                leLibelle = x[1]
                myData = wx.TreeItemData([leClasseur, leLibelle])
                child1 = self.arbre.AppendItem(self.root, leLibelle, self.imClasseur, data=myData)
                req = "SELECT dossier, libelle FROM dossiers WHERE classeur = %s ORDER BY dossier"%leClasseur
                self.c.execute(req)
                listeDossiers = self.c.fetchall()

                for y in listeDossiers:
                    leDossier = y[0]
                    leLibelle = y[1]
                    myData = wx.TreeItemData([leClasseur, leDossier, leLibelle])
                    child2 = self.arbre.AppendItem(child1, leLibelle, self.imDossier, data=myData)
                    req = "SELECT chemise, libelle FROM chemises where classeur = %s AND dossier = %s ORDER BY chemise"%(leClasseur, leDossier)
                    self.c.execute(req)
                    listeChemises = self.c.fetchall()

                    for z in listeChemises:
                        laChemise = z[0]
                        leLibelle = z[1]
                        myData = wx.TreeItemData([leClasseur, leDossier, laChemise, leLibelle])
                        child3 = self.arbre.AppendItem(child2, leLibelle, self.imChemise, data=myData)
            self.arbre.Expand(self.root)

    def ClickDroit(self, event):
        pt = event.GetPoint()
        item = event.GetItem()
        self.itemChoisi = item
        if (self.arbre.GetItemImage(item) == self.imRoot) :
            self.PopupMenu(self.popMenuRoot, pt)
        elif (self.arbre.GetItemImage(item) == self.imClasseur):
            self.PopupMenu(self.popMenuClasseur, pt)
        elif (self.arbre.GetItemImage(item) == self.imDossier):
            self.PopupMenu(self.popMenuDossier, pt)
        else :
            self.PopupMenu(self.popMenuChemise, pt)

    def AjoutClasseur(self, event):
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app, u"Saisir le nom du classeur à créer", "Nouveau classeur")
        val = dlgTxt.ShowModal()
        nom = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val==wx.ID_OK and nom != "":
            nom = "''".join(dlgTxt.GetValue().split("'"))
            nom = eval('u"%s"'%nom)
            req = "INSERT INTO classeurs(libelle) VALUES('%s')"%nom
            self.c.execute(req)
            res = self.c.execute("SELECT MAX(classeur) FROM classeurs")
            nouvClasseur = self.c.fetchone()[0]
            req = "SELECT libelle FROM classeurs WHERE classeur = %s"%nouvClasseur
            self.c.execute(req)
            res=self.c.fetchone()[0]
            myData = wx.TreeItemData([nouvClasseur, res])
            child1 = self.arbre.AppendItem(self.root, res, self.imClasseur, data=myData)
            self.arbre.Expand(self.itemChoisi)

    def AjoutDossier(self, event):
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app, u"Saisir le nom du dossier à créer", "Nouveau dossier")
        val = dlgTxt.ShowModal()
        resu = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val==wx.ID_OK and resu != "":
            nom = "''".join(resu.split("'"))
            nom = eval('u"%s"'%nom)
            leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
            req = "INSERT INTO dossiers(classeur, libelle) VALUES(%s, '%s')"%(leClasseur, nom)
            self.c.execute(req)
            res = self.c.execute("SELECT MAX(dossier) FROM dossiers")
            nouvDossier = self.c.fetchone()[0]
            myData = wx.TreeItemData([leClasseur, nouvDossier, resu])
            child1 = self.arbre.AppendItem(self.itemChoisi, resu, self.imDossier, data=myData)
            self.arbre.Expand(self.itemChoisi)

    def RenommeClasseur(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app,
                                    u"Saisir le nouveau nom du classeur",
                                    u"Renommer un classeur",
                                    defaultValue = self.arbre.GetPyData(self.itemChoisi)[1])
        val = dlgTxt.ShowModal()
        nom = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val==wx.ID_OK and nom != "":
            texte = "''".join(nom.split("'"))
            texte = eval('u"%s"'%nom)
            req = "UPDATE classeurs SET libelle = '%s' WHERE classeur = %s"%(texte, leClasseur)
            self.c.execute(req)
            self.arbre.SetItemText(self.itemChoisi, nom)

    def SupprimeClasseur(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        req = "SELECT COUNT(*) FROM documents WHERE classeur = %s"%leClasseur
        self.c.execute(req)
        nbre = int(self.c.fetchone()[0])
        if nbre > 0 :
            if nbre == 1 :
                phrase = u"Attention ! Vous allez avoir un document à reclasser après cette opération.\nVoulez-vous continuer ?"
            else:
                phrase = u"Attention ! vous allez avoir %s documents à reclasser après cette opération.\nVoulez-vous continuer ?"%nbre
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=phrase,
                                   caption=u"Suppression d'un classeur",
                                   style=wx.YES_NO|wx.ICON_EXCLAMATION)
        else:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"Voulez-vous vraiment supprimer ce classeur ?",
                                   caption=u"Suppression d'un classeur",
                                   style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES :
            if nbre > 0 :
                req = "UPDATE documents SET classeur = 0, dossier = 0, chemise = 0 WHERE classeur = %s"%leClasseur
                self.c.execute(req)
            req = "DELETE FROM classeurs WHERE classeur = %s"%leClasseur
            self.c.execute(req)
            req = "DELETE FROM dossiers WHERE classeur = %s"%leClasseur
            self.c.execute(req)
            req = "DELETE FROM chemises WHERE classeur = %s"%leClasseur
            self.c.execute(req)
            self.arbre.DeleteAllItems()
            self.Remplir()

    def AjoutChemise(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        leDossier = self.arbre.GetPyData(self.itemChoisi)[1]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app, u"Saisir le nom de la chemise à créer", "Nouvelle chemise")
        val = dlgTxt.ShowModal()
        resu = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val == wx.ID_OK and resu != "":
            nom = "''".join(resu.split("'"))
            nom = eval('u"%s"'%nom)
            req = "INSERT INTO chemises(classeur, dossier, libelle) VALUES(%s, %s, '%s')"%(leClasseur, leDossier, nom)
            self.c.execute(req)
            res = self.c.execute("SELECT MAX(chemise) FROM chemises")
            nouvChemise = self.c.fetchone()[0]
            myData = wx.TreeItemData([leClasseur, leDossier, nouvChemise, resu])
            child1 = self.arbre.AppendItem(self.itemChoisi, resu, self.imChemise, data=myData)
            self.arbre.Expand(self.itemChoisi)

    def RenommeDossier(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        leDossier = self.arbre.GetPyData(self.itemChoisi)[1]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app,
                                    u"Saisir le nouveau nom du dossier",
                                    "Renommer un dossier",
                                    defaultValue = self.arbre.GetPyData(self.itemChoisi)[2])
        val = dlgTxt.ShowModal()
        nom = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val==wx.ID_OK and nom != "":
            texte = "''".join(nom.split("'"))
            texte = eval('u"%s"'%nom)
            req = "UPDATE dossiers SET libelle = '%s' WHERE classeur = %s AND dossier = %s"%(texte, leClasseur, leDossier)
            self.c.execute(req)
            self.arbre.SetItemText(self.itemChoisi, nom)

    def SupprimeDossier(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        leDossier = self.arbre.GetPyData(self.itemChoisi)[1]
        req = "SELECT COUNT(*) FROM documents WHERE classeur = %s AND dossier = %s"%(leClasseur, leDossier)
        self.c.execute(req)
        nbre = int(self.c.fetchone()[0])
        if nbre > 0 :
            if nbre == 1 :
                phrase = u"Attention ! Vous allez avoir un document à reclasser après cette opération.\nVoulez-vous continuer ?"
            else:
                phrase = u"Attention ! vous allez avoir %s documents à reclasser après cette opération.\nVoulez-vous continuer ?"%nbre
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=phrase,
                                   caption=u"Suppression d'un dossier",
                                   style=wx.YES_NO|wx.ICON_EXCLAMATION)
        else:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"Voulez-vous vraiment supprimer ce dossier ?",
                                   caption=u"Suppression d'un dossier",
                                   style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES :
            if nbre > 0 :
                req = "UPDATE documents SET classeur = 0, dossier = 0, chemise = 0 WHERE classeur = %s AND dossier = %s"%(leClasseur, leDossier)
                self.c.execute(req)
            req = "DELETE FROM dossiers WHERE classeur = %s AND dossier = %s"%(leClasseur, leDossier)
            self.c.execute(req)
            req = "DELETE FROM chemises WHERE classeur = %s AND dossier = %s"%(leClasseur, leDossier)
            self.c.execute(req)
            self.arbre.DeleteAllItems()
            self.Remplir()

    def RenommeChemise(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        leDossier = self.arbre.GetPyData(self.itemChoisi)[1]
        laChemise = self.arbre.GetPyData(self.itemChoisi)[2]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app,
                                    u"Saisir le nouveau nom de la chemise",
                                    "Renommer une chemise",
                                    defaultValue = self.arbre.GetPyData(self.itemChoisi)[3])
        val = dlgTxt.ShowModal()
        nom = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val==wx.ID_OK and nom != "":
            texte = "''".join(nom.split("'"))
            texte = eval('u"%s"'%nom)
            req = "UPDATE chemises SET libelle = '%s' WHERE classeur = %s AND dossier = %s AND chemise = %s"%(texte, leClasseur, leDossier, laChemise)
            self.c.execute(req)
            self.arbre.SetItemText(self.itemChoisi, nom)

    def SupprimeChemise(self, event):
        leClasseur = self.arbre.GetPyData(self.itemChoisi)[0]
        leDossier = self.arbre.GetPyData(self.itemChoisi)[1]
        laChemise = self.arbre.GetPyData(self.itemChoisi)[2]
        req = "SELECT COUNT(*) FROM documents WHERE classeur = %s AND dossier = %s AND chemise = %s"%(leClasseur, leDossier, laChemise)
        self.c.execute(req)
        nbre = int(self.c.fetchone()[0])
        if nbre > 0 :
            if nbre == 1 :
                phrase = u"Attention ! Vous allez avoir un document à reclasser après cette opération.\nVoulez-vous continuer ?"
            else:
                phrase = u"Attention ! vous allez avoir %s documents à reclasser après cette opération.\nVoulez-vous continuer ?"%nbre
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=phrase,
                                   caption=u"Suppression d'une chemise",
                                   style=wx.YES_NO|wx.ICON_EXCLAMATION)
        else:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"Voulez-vous vraiment supprimer cette chemise ?",
                                   caption=u"Suppression d'une chemise",
                                   style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES :
            if nbre > 0 :
                req = "UPDATE documents SET classeur = 0, dossier = 0, chemise = 0 WHERE classeur = %s AND dossier = %s AND chemise = %s"%(leClasseur, leDossier, laChemise)
                self.c.execute(req)
            req = "DELETE FROM chemises WHERE classeur = %s AND dossier = %s AND chemise = %s"%(leClasseur, leDossier, laChemise)
            self.c.execute(req)
            self.arbre.DeleteAllItems()
            self.Remplir()

class MonAffichTexte(wx.ScrolledWindow):
    def __init__(self, parent, fichier):
        wx.ScrolledWindow.__init__(self, parent=parent, id=-1)
        sizer = wx.BoxSizer(wx.VERTICAL)
        fic = open(fichier, "r")
        contenu = fic.read()
        fic.close()
        self.visuTexte = wx.TextCtrl(self,
                                     id=-1,
                                     value=contenu,
                                     style=wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_LEFT|wx.TE_DONTWRAP)
        sizer.Add(self.visuTexte, 1, wx.EXPAND)
        self.SetSizer(sizer)
        sizer.Fit(self)
        self.SetScrollRate(20, 20)

class APropos(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, parent=GLOBVAR.app, title=u"A propos de Pap'rass", size=(500, 500))
        image = wx.StaticBitmap(self, -1,
                                wx.Bitmap(os.path.join(GLOBVAR.themedir, "hautdepage.png"),
                                wx.BITMAP_TYPE_PNG))
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(image, 0, flag=wx.EXPAND|wx.TOP|wx.BOTTOM, border=10)
        ntbook = wx.Notebook(self, -1)
        copyright = wx.Panel(ntbook, -1)
        copSizer= wx.BoxSizer(wx.VERTICAL)
        mess = u"Pap'rass 2.06\n"
        mess = mess + u"réalisé en wxPython\n\n"
        mess = mess + u"copyright (C) 2007-2008 Alain DELGRANGE aka bipede\n"
        mess = mess + u"Licence GNU-GPL version 2\n\n"
        mess = mess + u"Numérisez, classez et retrouvez vos paperasses facilement."
        label = wx.StaticText(copyright, -1, label=mess, style=wx.ALIGN_CENTRE)
        copSizer.AddStretchSpacer()
        copSizer.Add(label, 1, wx.EXPAND)
        copSizer.AddStretchSpacer()
        copyright.SetSizer(copSizer)
        copSizer.Fit(copyright)
        copyright.SetAutoLayout(True)
        ntbook.AddPage(copyright, u"Copyright")
        licence = MonAffichTexte(ntbook, os.path.join(GLOBVAR.appdir, "licence", "GNU_GENERAL_LICENCE.txt"))
        ntbook.AddPage(licence, u"licence")
        sizer.Add(ntbook, 1, flag=wx.EXPAND)
        sizerbouton = self.CreateButtonSizer(wx.OK)
        sizer.Add(sizerbouton, flag=wx.ALIGN_RIGHT|wx.ALL, border=10)
        self.SetSizer(sizer)
        self.SetAutoLayout(True)

class RechercheGlobale(wx.ScrolledWindow):
    def __init__(self, parent, origine):
        wx.ScrolledWindow.__init__(self, parent=parent, id=-1, style=wx.BORDER_SUNKEN)
        self.itemChoisi = None
        self.origine= origine
        fenSizer=wx.BoxSizer(wx.VERTICAL)
        self.listeImage = wx.ImageList(24, 24)
        self.imRoot = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "racine.png"), wx.BITMAP_TYPE_PNG))
        self.imClasseur = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "classeur.png"), wx.BITMAP_TYPE_PNG))
        self.imDossier = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "dossier.png"), wx.BITMAP_TYPE_PNG))
        self.imChemise = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "chemise.png"), wx.BITMAP_TYPE_PNG))
        self.imDocument = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "document.png"), wx.BITMAP_TYPE_PNG))
        self.arbre = wx.TreeCtrl(self, -1)
        fenSizer.Add(self.arbre, 1, wx.EXPAND)
        self.SetSizer(fenSizer)
        self.SetAutoLayout(True)
        self.arbre.SetImageList(self.listeImage)
        self.c = GLOBVAR.base.cursor()
        self.Remplir()

        self.popMenuDoc= wx.Menu()
        self.popMenuDoc.Append(wx.ID_FILE1, u"Visualiser le document")
        self.popMenuDoc.Append(wx.ID_FILE2, u"Renommer le document")
        self.popMenuDoc.Append(wx.ID_FILE3, u"Exporter le document")
        self.popMenuDoc.Append(wx.ID_FILE4, u"Remettre le document à classer")
        self.popMenuDoc.Append(wx.ID_FILE5, u"Supprimer le document")

        self.arbre.Bind(wx.EVT_TREE_ITEM_RIGHT_CLICK, self.ClickDroit)
        self.arbre.Bind(wx.EVT_TREE_SEL_CHANGED, self.ClickGauche)

        self.Bind(wx.EVT_MENU, self.Visualiser, id = wx.ID_FILE1)
        self.Bind(wx.EVT_MENU, self.Renommer, id = wx.ID_FILE2)
        self.Bind(wx.EVT_MENU, self.Exporter, id = wx.ID_FILE3)
        self.Bind(wx.EVT_MENU, self.Declasser, id = wx.ID_FILE4)
        self.Bind(wx.EVT_MENU, self.Supprimer, id = wx.ID_FILE5)

    def ReinitArbre(self):
        self.arbre.DeleteAllItems()
        self.Remplir()

    def Remplir(self):
        self.root = self.arbre.AddRoot(u"Plan de classement des documents", self.imRoot)
        self.itemChoisi = self.root

        self.c.execute("SELECT COUNT(*) FROM classeurs")
        if (self.c.fetchall()[0][0] > 0):
            self.c.execute("SELECT classeur, libelle FROM classeurs ORDER BY majuscules(libelle)")
            listeClasseurs = self.c.fetchall()

            for x in listeClasseurs:
                leClasseur = x[0]
                leLibelle = x[1]
                child1 = self.arbre.AppendItem(self.root, leLibelle, self.imClasseur)
                req = "SELECT dossier, libelle FROM dossiers WHERE classeur = %s ORDER BY majuscules(libelle)"%leClasseur
                self.c.execute(req)
                listeDossiers = self.c.fetchall()

                for y in listeDossiers:
                    leDossier = y[0]
                    leLibelle = y[1]
                    child2 = self.arbre.AppendItem(child1, leLibelle, self.imDossier)
                    req = "SELECT chemise, libelle FROM chemises where classeur = %s AND dossier = %s ORDER BY majuscules(libelle)"%(leClasseur, leDossier)
                    self.c.execute(req)
                    listeChemises = self.c.fetchall()

                    for z in listeChemises:
                        laChemise = z[0]
                        leLibelle = z[1]
                        child3 = self.arbre.AppendItem(child2, leLibelle, self.imChemise)
                        req = "SELECT enreg, titre , date FROM documents WHERE classeur = %s AND dossier = %s AND chemise = %s ORDER BY date"%(leClasseur, leDossier, laChemise)
                        self.c.execute(req)
                        listeDoc = self.c.fetchall()
                        if len(listeDoc )> 0:
                            for a in listeDoc:
                                annee = a[2].split("-")[0]
                                mois = a[2].split("-")[1]
                                jour = a[2].split("-")[2]
                                laDate = jour + "/" + mois + "/" + annee
                                leLibelle = laDate + " " + a[1]
                                myData = wx.TreeItemData(a)
                                child4 = self.arbre.AppendItem(child3, leLibelle, self.imDocument, data=myData)
                                self.arbre.SetItemTextColour(child4, wx.BLUE)

            self.arbre.Expand(self.root)

    def ClickGauche(self, event):
        self.origine.SetImage()

    def ClickDroit(self, event):
        pt = event.GetPoint()
        item = event.GetItem()
        self.itemChoisi = item
        if (self.arbre.GetItemImage(item) == self.imDocument) :
            self.PopupMenu(self.popMenuDoc, pt)

    def Declasser(self, event):
        enreg = self.arbre.GetPyData(self.itemChoisi)[0]
        titre = self.arbre.GetPyData(self.itemChoisi)[1]
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               message = u"Voulez-vous vraiment remettre ce document\nintitulé \"%s\"\nà classer ?"% titre,
                               caption=u"Remettre un document à classer",
                               style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES:
            req = "UPDATE documents SET classeur = 0, dossier = 0, chemise = 0 WHERE enreg = %s"%enreg
            self.c.execute(req)
            self.arbre.Delete(self.itemChoisi)

    def Renommer(self, event):
        enreg = self.arbre.GetPyData(self.itemChoisi)[0]
        titre = self.arbre.GetPyData(self.itemChoisi)[1]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app,
                                    u"Saisir le nouveau titre de ce document",
                                    u"Renommer le document",
                                    defaultValue = titre)
        val = dlgTxt.ShowModal()
        newtitre = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val == wx.ID_OK:
            if newtitre != "":
                resu = "''".join(newtitre.split("'"))
                resu = eval('u"%s"'%resu)
                req = "UPDATE documents SET titre = '%s' WHERE enreg = %s"%(resu, enreg)
                self.c.execute(req)
                self.arbre.SetItemText(self.itemChoisi, newtitre)
            else:
                dlg = MessageDialog(parent = GLOBVAR.app,
                                    message = u"Vous devez donner un titre au document",
                                    caption = u"Opération impossible",
                                    style = wx.OK|wx.ICON_ERROR)
                val = dlg.ShowModal()
                dlg.Destroy()

    def Visualiser(self, event):
        libdate = self.arbre.GetPyData(self.itemChoisi)[2]
        titre = self.arbre.GetPyData(self.itemChoisi)[1]
        enreg = self.arbre.GetPyData(self.itemChoisi)[0]
        mois = libdate.split("-")[1]
        annee = libdate.split("-")[0]
        req = "SELECT nbpages FROM documents WHERE enreg = %s "%(enreg)
        self.c.execute(req)
        pages = self.c.fetchone()[0]
        racine = os.path.join(GLOBVAR.docdir, annee, mois)
        if pages == 1:
            fichier = "%s-1"%enreg
            if os.path.isfile(os.path.join(racine,  fichier + ".txt")):
                fic = os.path.join(racine,  fichier + ".txt")
                dlg = AffichageTextesDialog(titre, fic)
                val = dlg.ShowModal()
                dlg.Destroy()
            elif IsItImage(os.path.join(racine,  fichier)):
                liste = []
                liste.append(IsItImage(os.path.join(racine,  fichier)))
                self.origine.SetImage(liste)
            elif IsItOoffice(os.path.join(racine,  fichier)):
                chemin = IsItOoffice(os.path.join(racine,  fichier))
                if WIN:
                    os.startfile(chemin)
                else:
                    commande= 'ooffice "%s"'%chemin
                    os.system(commande)
            else:
                fichier = fichier + ".pdf"
                if WIN:
                    os.startfile(os.path.join(racine,  fichier))
                else:
                    commande= '%s "%s"'%(GLOBVAR.visupdf, os.path.join(racine,  fichier))
                    os.system(commande)
        else:
            maListe = []
            fic = str(enreg) + "-"
            for x in range(pages):
                fichier = IsItImage(os.path.join(racine,  fic + str(x+1)))
                maListe.append(fichier)
            self.origine.SetImage(maListe)
            
    def Exporter(self, event):
        libdate = self.arbre.GetPyData(self.itemChoisi)[2]
        titre = self.arbre.GetPyData(self.itemChoisi)[1]
        enreg = self.arbre.GetPyData(self.itemChoisi)[0]
        mois = libdate.split("-")[1]
        annee = libdate.split("-")[0]
        req = "SELECT nbpages FROM documents WHERE enreg = %s"%(enreg)
        self.c.execute(req)
        pages = self.c.fetchone()[0]
        racine = os.path.join(GLOBVAR.docdir, annee, mois)
        fic = str(enreg) + "-"
        ooo = False
        for term in GLOBVAR.listeoo:
            fin = term.lower()
            if os.path.isfile(os.path.join(racine, fic + "1." + fin)):
                fichier = os.path.join(racine, fic + "1." + fin)
                ooo = True
                terminaison = fin
        if ooo:
            parDefaut = "sans titre.%s"%terminaison
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document Open-Office",
                                defaultFile=parDefaut,
                                wildcard="*.%s"%terminaison,
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)

        elif os.path.isfile(os.path.join(racine, fic + "1.pdf")):
            fichier = os.path.join(racine, fic + "1.pdf")
            leFichier = "sans titre.pdf"
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document PDF",
                                defaultFile=leFichier,
                                wildcard="*.pdf",
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)

        elif os.path.isfile(os.path.join(racine, fic + "1.txt")):
            fichier = os.path.join(racine, fic + "1.txt")
            leFichier = "sans titre.txt"
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document texte",
                                defaultFile=leFichier,
                                wildcard="*.txt",
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)
        else:
            if pages > 1:
                dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                       message=u"Ce document est composé de %s pages.\nVoulez-vous continuer ?"% pages,
                                       caption=u"Exporter un document",
                                       style=wx.YES_NO|wx.ICON_QUESTION)
                val = dlg.ShowModal()
                dlg.Destroy()
            else:
                val = wx.ID_YES
            if val == wx.ID_YES:
                for x in range(pages):
                    doc = IsItImage(os.path.join(racine, fic + str(x + 1)))
                    eclate = doc.split(".")
                    term = eclate[len(eclate)-1]
                    leFichier = "image-%s.%s"%(x+1, term)
                    card = "*.%s"%term
                    dlg = wx.FileDialog(parent=GLOBVAR.app,
                                        message=u"Exporter l'image n° %s"%(x+1),
                                        defaultFile=leFichier,
                                        wildcard=card,
                                        style= wx.FD_SAVE)
                    rep = dlg.ShowModal()
                    ficsauve = dlg.GetPath()
                    dlg.Destroy()
                    if rep == wx.ID_OK:
                        shutil.copyfile(doc, ficsauve)

    def Supprimer(self, event):
        libdate = self.arbre.GetPyData(self.itemChoisi)[2]
        titre = self.arbre.GetPyData(self.itemChoisi)[1]
        enreg = self.arbre.GetPyData(self.itemChoisi)[0]
        mois = libdate.split("-")[1]
        annee = libdate.split("-")[0]
        req = "SELECT nbpages FROM documents WHERE enreg = %s"%(enreg)
        self.c.execute(req)
        pages = self.c.fetchall()[0][0]
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               message=u"Voulez-vous vraiment détruire ce document\nintitulé \"%s\" ?"% titre,
                               caption=u"Suppression d'un document",
                               style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES:
            racine = os.path.join(GLOBVAR.docdir, annee, mois)
            fic = str(enreg) +"-"
            ooo = False
            for term in GLOBVAR.listeoo:
                fin = term.lower()
                if os.path.isfile(os.path.join(racine, fic + "1." + fin)):
                    fichier = os.path.join(racine, fic + "1." + fin)
                    ooo = True
                    terminaison = fin
            if ooo:
                os.remove(os.path.join(racine, fic + "1." + terminaison))
            elif os.path.isfile(os.path.join(racine, fic + "1.pdf")):
                os.remove(os.path.join(racine, fic + "1.pdf"))
            elif os.path.isfile(os.path.join(racine, fic + "1.txt")):
                os.remove(os.path.join(racine, fic + "1.txt"))
            else:
                for x in range(pages):
                    doc = os.path.join(racine, fic + str(x + 1))
                    os.remove(IsItImage(doc))
            req = "DELETE FROM documents WHERE enreg = %s"%(enreg)
            self.c.execute(req)
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                message=u"le document a été supprimé",
                                caption=u"Suppression",
                                style=wx.OK|wx.ICON_INFORMATION)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.arbre.Delete(self.itemChoisi)
            self.origine.ReinitCle()


class RechercheMotCle(wx.Panel):
    def __init__(self, parent, origine):
        wx.Panel.__init__(self, parent=parent, id=-1, style=wx.BORDER_SUNKEN)
        self.origine = origine
        self.item = None
        self.resultat=[]
        boxPrinc = wx.BoxSizer(wx.VERTICAL)
        self.c = GLOBVAR.base.cursor()
        etiquette = wx.StaticText(self, -1, label=u"Recherche sur un mot clé", style=wx.CENTRE)
        boxPrinc.Add(etiquette,flag=wx.CENTRE|wx.ALL, border=5)
        box = wx.BoxSizer(wx.HORIZONTAL)
        self.saisie = wx.TextCtrl(self, -1, style=wx.TE_PROCESS_ENTER)
        boxPrinc.Add(self.saisie, flag=wx.EXPAND|wx.ALL, border=5)
        self.bouton1 = wx.Button(self, -1, u"Rechercher")
        self.Bind(wx.EVT_BUTTON, self.Rechercher, self.bouton1)
        self.bouton2 = wx.Button(self, -1, u"Effacer")
        self.Bind(wx.EVT_BUTTON, self.Effacer, self.bouton2)
        self.bouton2.Enable(False)
        box.Add((0,0), 1)
        box.Add(self.bouton1, 5, flag=wx.EXPAND)
        box.Add((0,0), 1)
        box.Add(self.bouton2, 5, flag=wx.EXPAND)
        box.Add((0,0), 1)
        boxPrinc.Add(box, flag=wx.EXPAND|wx.ALL, border=5)
        self.grille = wx.ScrolledWindow(self, -1)
        boxGrid = wx.BoxSizer(wx.VERTICAL)
        self.listResultat = wx.ListCtrl(self.grille, -1, style=wx.LC_REPORT|wx.LC_SINGLE_SEL)
        self.listResultat.InsertColumn(0, u"Date")
        self.listResultat.InsertColumn(1, u"Titre du document")
        boxGrid.Add(self.listResultat, 1, wx.EXPAND)
        self.grille.SetSizer(boxGrid)
        self.grille.SetAutoLayout(True)
        boxPrinc.Add(self.grille, 1, flag=wx.EXPAND|wx.ALL, border=5)
        self.SetSizer(boxPrinc)
        boxPrinc.Fit(self)
        self.SetAutoLayout(True)
        self.popMenuDoc= wx.Menu()
        self.popMenuDoc.Append(wx.ID_FILE1, u"Visualiser le document")
        self.popMenuDoc.Append(wx.ID_FILE2, u"Renommer le document")
        self.popMenuDoc.Append(wx.ID_FILE3, u"Exporter le document")
        self.popMenuDoc.Append(wx.ID_FILE4, u"Remettre le document à classer")
        self.popMenuDoc.Append(wx.ID_FILE5, u"Supprimer le document")

        wx.EVT_SIZE(self, self.OnSize)
        self.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, self.OnRightClick, self.listResultat)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnLeftClick, self.listResultat)
        self.saisie.Bind(wx.EVT_TEXT_ENTER, self.Rechercher)
        self.Bind(wx.EVT_MENU, self.Visualiser, id = wx.ID_FILE1)
        self.Bind(wx.EVT_MENU, self.Renommer, id = wx.ID_FILE2)
        self.Bind(wx.EVT_MENU, self.Exporter, id = wx.ID_FILE3)
        self.Bind(wx.EVT_MENU, self.Declasser, id = wx.ID_FILE4)
        self.Bind(wx.EVT_MENU, self.Supprimer, id = wx.ID_FILE5)

    def OnLeftClick(self, event):
        self.origine.SetImage()

    def OnRightClick(self, event):
        self.item, flag = self.listResultat.HitTest(event.GetPosition())
        if self.item > -1:
            pt = event.GetPoint()
            self.listResultat.PopupMenu(self.popMenuDoc, pt)
        else:
            self.item = None

    def OnSize(self, event):
        self.Layout()
        larg, haut = self.grille.GetClientSizeTuple()
        larg = larg
        if self.listResultat.GetItemCount() == 0:
            self.listResultat.SetColumnWidth(0, int(larg/4))
            self.listResultat.SetColumnWidth(1, int((larg/4)*3))
        else:
            self.listResultat.SetColumnWidth(0, wx.LIST_AUTOSIZE)
            self.listResultat.SetColumnWidth(1, wx.LIST_AUTOSIZE)
            larg1 = self.listResultat.GetColumnWidth(0)
            larg2 = self.listResultat.GetColumnWidth(1)
            if (larg1 + larg2) < larg:
                self.listResultat.SetColumnWidth(1, larg-larg1)

    def Effacer(self, event=None):
        self.saisie.SetValue("")
        self.resultat=[]
        if self.listResultat.GetItemCount() > 0:
            self.listResultat.DeleteAllItems()
        self.origine.SetImage()
        self.bouton1.Enable(True)
        self.bouton2.Enable(False)
        self.SendSizeEvent()
        self.saisie.SetFocus()

    def Rechercher(self, event=None):
        self.bouton1.Enable(False)
        self.bouton2.Enable(True)
        if self.listResultat.GetItemCount() > 0:
            self.listResultat.DeleteAllItems()
        mot = self.saisie.GetValue()
        if mot != "":
            mot = eval('u"%s"'%mot)
            req = "SELECT annee, mois, enreg, date, titre, nbpages FROM documents"
            self.c.execute(req)
            liste = self.c.fetchall()
            if len(liste) > 0:
                self.resultat = []
                for x in liste:
                    titre = x[4]
                    if mot.upper() in titre.upper():
                        self.resultat.append(x)
                if len(self.resultat) > 0:
                    for x in range(len(self.resultat)):
                        eclate = self.resultat[x][3].split("-")
                        date = eclate[2] + "/" + eclate[1] + "/" + eclate[0]
                        index = self.listResultat.InsertStringItem(x, date)
                        self.listResultat.SetStringItem(index, 0, date)
                        self.listResultat.SetStringItem(index, 1, self.resultat[x][4])
                    self.SendSizeEvent()
                    return
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                caption = u"Recherche non aboutie".encode('utf-8'),
                                message=u"Aucun titre de document ne contient le mot clé indiqué",
                                style = wx.OK|wx.ICON_INFORMATION)
            val = dlg.ShowModal()
            dlg.Destroy()

    def Declasser(self, event):
        enreg = self.resultat[self.item][2]
        titre = self.resultat[self.item][4]
        req = "SELECT classeur FROM documents WHERE enreg = %s"%enreg
        self.c.execute(req)
        if self.c.fetchone()[0] == 0:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message = u"Le document désigné n'est pas classé",
                                   caption=u"Opération inutile",
                                   style=wx.OK|wx.ICON_ERROR)
            val = dlg.ShowModal()
            dlg.Destroy()
        else:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message = u"Voulez-vous vraiment remettre ce document\nintitulé \"%s\"\nà classer ?"% titre,
                                   caption=u"Remettre un document à classer",
                                   style=wx.YES_NO|wx.ICON_QUESTION)
            val = dlg.ShowModal()
            dlg.Destroy()
            if val == wx.ID_YES:
                req = "UPDATE documents SET classeur = 0, dossier = 0, chemise = 0 WHERE enreg = %s"%enreg
                self.c.execute(req)
                self.origine.ReinitGlobal()
                dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                       message = u"Le document a été remis à classer ?",
                                       caption=u"Opération réalisée",
                                       style=wx.OK|wx.ICON_INFORMATION)
                val = dlg.ShowModal()
                dlg.Destroy()

    def Renommer(self, event):
        enreg = self.resultat[self.item][2]
        titre = self.resultat[self.item][4]
        dlgTxt = wx.TextEntryDialog(GLOBVAR.app,
                                    u"Saisir le nouveau titre de ce document",
                                    u"Renommer le document",
                                    defaultValue = titre)
        val = dlgTxt.ShowModal()
        newtitre = dlgTxt.GetValue()
        dlgTxt.Destroy()
        if val == wx.ID_OK:
            if newtitre != "":
                resu = "''".join(newtitre.split("'"))
                resu = eval('u"%s"'%resu)
                req = "UPDATE documents SET titre = '%s' WHERE enreg = %s"%(resu, enreg)
                self.c.execute(req)
                self.listResultat.SetStringItem(self.item, 1, newtitre)
                self.origine.ReinitGlobal()
            else:
                dlg = MessageDialog(parent = GLOBVAR.app,
                                    message = u"Vous devez donner un titre au document",
                                    caption = u"Opération impossible",
                                    style = wx.OK|wx.ICON_ERROR)
                val = dlg.ShowModal()
                dlg.Destroy()

    def Visualiser(self, event):
        titre = self.resultat[self.item][4]
        enreg = self.resultat[self.item][2]
        mois = self.resultat[self.item][1]
        annee = self.resultat[self.item][0]
        pages = self.resultat[self.item][5]
        racine = os.path.join(GLOBVAR.docdir, annee, mois)
        if pages == 1:
            fichier = "%s-1"%enreg
            if os.path.isfile(os.path.join(racine,  fichier + ".txt")):
                fic = os.path.join(racine,  fichier + ".txt")
                dlg = AffichageTextesDialog(titre, fic)
                val = dlg.ShowModal()
                dlg.Destroy()
            elif IsItImage(os.path.join(racine,  fichier)):
                liste = []
                liste.append(IsItImage(os.path.join(racine,  fichier)))
                self.origine.SetImage(liste)
            elif IsItOoffice(os.path.join(racine,  fichier)):
                chemin = IsItOoffice(os.path.join(racine,  fichier))
                if WIN:
                    os.startfile(chemin)
                else:
                    commande= 'ooffice "%s"'%chemin
                    os.system(commande)
            else:
                fichier = fichier + ".pdf"
                if WIN:
                    os.startfile(os.path.join(racine,  fichier))
                else:
                    commande= '%s "%s"'%(GLOBVAR.visupdf, os.path.join(racine,  fichier))
                    os.system(commande)
        else:
            maListe = []
            fic = str(enreg) + "-"
            for x in range(pages):
                fichier = IsItImage(os.path.join(racine,  fic + str(x+1)))
                maListe.append(fichier)
            self.origine.SetImage(maListe)
            
    def Exporter(self, event):
        titre = self.resultat[self.item][4]
        enreg = self.resultat[self.item][2]
        mois = self.resultat[self.item][1]
        annee = self.resultat[self.item][0]
        pages = self.resultat[self.item][5]
        racine = os.path.join(GLOBVAR.docdir, annee, mois)
        fic = str(enreg) + "-"
        ooo = False
        for term in GLOBVAR.listeoo:
            fin = term.lower()
            if os.path.isfile(os.path.join(racine, fic + "1." + fin)):
                fichier = os.path.join(racine, fic + "1." + fin)
                ooo = True
                terminaison = fin
        if ooo:
            parDefaut = "sans titre.%s"%terminaison
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document Open-Office",
                                defaultFile=parDefaut,
                                wildcard="*.%s"%terminaison,
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)

        elif os.path.isfile(os.path.join(racine, fic + "1.pdf")):
            fichier = os.path.join(racine, fic + "1.pdf")
            leFichier = "sans titre.pdf"
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document PDF",
                                defaultFile=leFichier,
                                wildcard="*.pdf",
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)

        elif os.path.isfile(os.path.join(racine, fic + "1.txt")):
            fichier = os.path.join(racine, fic + "1.txt")
            leFichier = "sans titre.txt"
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=u"Exporter un document texte",
                                defaultFile=leFichier,
                                wildcard="*.txt",
                                style= wx.FD_SAVE)
            rep = dlg.ShowModal()
            ficsauve = dlg.GetPath()
            dlg.Destroy()
            if rep == wx.ID_OK:
                shutil.copyfile(fichier, ficsauve)
        else:
            if pages > 1:
                dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                       message=u"Ce document est composé de %s pages.\nVoulez-vous continuer ?"% pages,
                                       caption=u"Exporter un document",
                                       style=wx.YES_NO|wx.ICON_QUESTION)
                val = dlg.ShowModal()
                dlg.Destroy()
            else:
                val = wx.ID_YES
            if val == wx.ID_YES:
                for x in range(pages):
                    doc = IsItImage(os.path.join(racine, fic + str(x + 1)))
                    eclate = doc.split(".")
                    term = eclate[len(eclate)-1]
                    leFichier = "image-%s.%s"%(x+1, term)
                    card = "*.%s"%term
                    dlg = wx.FileDialog(parent=GLOBVAR.app,
                                        message=u"Exporter l'image n° %s"%(x+1),
                                        defaultFile=leFichier,
                                        wildcard=card,
                                        style= wx.FD_SAVE)
                    rep = dlg.ShowModal()
                    ficsauve = dlg.GetPath()
                    dlg.Destroy()
                    if rep == wx.ID_OK:
                        shutil.copyfile(doc, ficsauve)

    def Supprimer(self, event):
        titre = self.resultat[self.item][4]
        enreg = self.resultat[self.item][2]
        mois = self.resultat[self.item][1]
        annee = self.resultat[self.item][0]
        pages = self.resultat[self.item][5]
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               message=u"Voulez-vous vraiment détruire ce document\nintitulé \"%s\" ?"% titre,
                               caption=u"Suppression d'un document",
                               style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES:
            racine = os.path.join(GLOBVAR.docdir, annee, mois)
            fic = str(enreg) +"-"
            ooo = False
            for term in GLOBVAR.listeoo:
                fin = term.lower()
                if os.path.isfile(os.path.join(racine, fic + "1." + fin)):
                    fichier = os.path.join(racine, fic + "1." + fin)
                    ooo = True
                    terminaison = fin
            if ooo:
                os.remove(os.path.join(racine, fic + "1." + terminaison))
            elif os.path.isfile(os.path.join(racine, fic + "1.pdf")):
                os.remove(os.path.join(racine, fic + "1.pdf"))
            elif os.path.isfile(os.path.join(racine, fic + "1.txt")):
                os.remove(os.path.join(racine, fic + "1.txt"))
            else:
                for x in range(pages):
                    doc = os.path.join(racine, fic + str(x + 1))
                    os.remove(IsItImage(doc))
            req = "DELETE FROM documents WHERE enreg = %s"%(enreg)
            self.c.execute(req)
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"le document a été supprimé",
                                   caption=u"Suppression",
                                   style=wx.OK|wx.ICON_INFORMATION)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.Rechercher()
            self.origine.ReinitGlobal()

class Recherche(wx.SplitterWindow):
    def __init__(self, parent):
        wx.SplitterWindow.__init__(self, parent = parent, id = -1, style=wx.SP_3D)
        l, h = parent.GetClientSizeTuple()
        larg = int(l/3)
        panel1=wx.Panel(self, -1)
        box1 = wx.BoxSizer(wx.VERTICAL)
        etiquette = wx.StaticText(panel1,
                                  id= -1,
                                  label = u"Vous accéderez aux menus en cliquant\n les items avec le bouton droit de la souris",
                                  style = wx.ALIGN_CENTRE)
        box1.Add(etiquette, flag= wx.CENTRE|wx.ALL, border=3)
        ntbook = wx.Notebook(panel1, -1)
        self.glob = RechercheGlobale(ntbook, self)
        ntbook.AddPage(self.glob, u"Par le plan de classement")
        self.cle = RechercheMotCle(ntbook, self)
        ntbook.AddPage(self.cle, u"A l'aide d'un mot clé")
        box1.Add(ntbook, 1, wx.EXPAND|wx.ALL, border=3)
        panel1.SetSizer(box1)
        box1.Fit(panel1)
        panel1.SetAutoLayout(True)
        panel2=wx.Panel(self, -1)
        box2 = wx.BoxSizer(wx.VERTICAL)
        self.ecran = Affichage(panel2)
        box2.Add(self.ecran, 1, wx.EXPAND|wx.ALL, border=3)
        panel2.SetSizer(box2)
        box2.Fit(panel2)
        panel2.SetAutoLayout(True)
        self.SetMinimumPaneSize(100)
        self.SplitVertically(panel1, panel2, larg)

        ntbook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGING, self.Bascule)

    def Bascule(self, event):
        self.SetImage()

    def SetImage(self, liste=None):
        if liste:
            self.ecran.SetDocument(liste)
        else:
            self.ecran.SetDocument()

    def ReinitGlobal(self):
        self.glob.ReinitArbre()

    def ReinitCle(self):
        self.cle.Effacer()

class MyPrintout(wx.Printout):
    def __init__(self, bmp):
        wx.Printout.__init__(self)
        self.bmp = bmp

    def OnBeginDocument(self, start, end):
        return super(MyPrintout, self).OnBeginDocument(start, end)

    def OnEndDocument(self):
        super(MyPrintout, self).OnEndDocument()

    def OnBeginPrinting(self):
        super(MyPrintout, self).OnBeginPrinting()

    def OnEndPrinting(self):
        super(MyPrintout, self).OnEndPrinting()

    def OnPreparePrinting(self):
        super(MyPrintout, self).OnPreparePrinting()

    def OnPrintPage(self, page):
        dc = self.GetDC()
        maxX = self.bmp.GetWidth()
        maxY = self.bmp.GetHeight()
        self.FitThisSizeToPage(wx.Size(maxX, maxY))
        dc.DrawBitmap(self.bmp, 0, 0)

        return True

class Affichage(wx.Panel):
    def __init__(self, parent, images=[]):
        wx.Panel.__init__(self, parent=parent, id=-1, style=wx.BORDER_SIMPLE)
        self.principale = parent
        self.images = images
        self.loupe = 0
        self.rotation = 0
        self.affichage = None
        self.image=None
        self.bmp=None
        self.box = wx.BoxSizer(wx.VERTICAL)
        if len(self.images)==0:
            texte = u"Page 0 sur 0"
            self.page=0
        else:
            texte = u"Page 1 sur %s"%len(self.images)
            self.page=1
        self.etiquette = wx.StaticText(self, -1, label=texte)
        self.box.Add(self.etiquette, flag=wx.CENTRE|wx.ALL, border=4)
        self.barre = wx.ToolBar(self, -1)

        self.barre.SetToolBitmapSize((32, 32))

        self.first_bouton = self.barre.AddSimpleTool(wx.ID_FILE1,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "premier.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Aller à la première page")
        self.back_bouton = self.barre.AddSimpleTool(wx.ID_FILE2,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "avant.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Aller à la page précédente")
        self.forward_bouton = self.barre.AddSimpleTool(wx.ID_FILE3,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "apres.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Aller à la page suivante")
        self.last_bouton = self.barre.AddSimpleTool(wx.ID_FILE4,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "dernier.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Aller à la dernière page")
        self.barre.AddSeparator()
        self.zoom_in_bouton = self.barre.AddSimpleTool(wx.ID_FILE5,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "plus.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Zoom avant")
        self.zoom_out_bouton = self.barre.AddSimpleTool(wx.ID_FILE6,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "moins.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Zoom arrière")
        self.zoom_100_bouton = self.barre.AddSimpleTool(wx.ID_FILE7,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "egal.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Taille initiale")
        self.zoom_fit_bouton = self.barre.AddSimpleTool(wx.ID_FILE8,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "ajuste.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"ajuster à la largeur de la fenêtre")
        self.barre.AddSeparator()
        self.rotate_bouton = self.barre.AddSimpleTool(wx.ID_FILE9,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "rotation.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Faire pivoter l'image")
        self.imprim_bouton = self.barre.AddSimpleTool(wx.ID_PRINT,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "imprimer.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Imprimer la page en cours")
        self.barre.AddSeparator()
        self.save_bouton = self.barre.AddSimpleTool(wx.ID_SAVE,
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "save.png"), wx.BITMAP_TYPE_PNG),
                        shortHelpString = u"Sauvegarder l'orientation de l'image")
        self.barre.Realize()
        self.box.Add(self.barre, flag=wx.EXPAND|wx.ALL, border=4)
        self.affichage=Apercu(self)
        self.box.Add(self.affichage, 1, flag=wx.EXPAND|wx.ALL, border=4)
        self.SetDocument(images)
        self.SetSizer(self.box)
        self.box.Fit(self)
        self.SetAutoLayout(True)
        wx.EVT_TOOL(self, wx.ID_FILE1, self.OnClickFirst)
        wx.EVT_TOOL(self, wx.ID_FILE2, self.OnClickBack)
        wx.EVT_TOOL(self, wx.ID_FILE3, self.OnClickForward)
        wx.EVT_TOOL(self, wx.ID_FILE4, self.OnClickLast)
        wx.EVT_TOOL(self, wx.ID_FILE5, self.ZoomerPlus)
        wx.EVT_TOOL(self, wx.ID_FILE6, self.ZoomerMoins)
        wx.EVT_TOOL(self, wx.ID_FILE7, self.Retablir)
        wx.EVT_TOOL(self, wx.ID_FILE8, self.Ajuster)
        wx.EVT_TOOL(self, wx.ID_FILE9, self.Rotate)
        wx.EVT_TOOL(self, wx.ID_PRINT, self.Print)
        wx.EVT_TOOL(self, wx.ID_SAVE, self.Save)

    def SetDocument(self, liste=[], page=1):
        self.images = liste
        if len(self.images) > 0:
            texte = u"Page %s sur %s"%(page, len(self.images))
            self.etiquette.SetLabel(texte)
            self.image = wx.Bitmap(self.images[page-1], wx.BITMAP_TYPE_ANY)
            self.affichage.SetImage(self.image)
            if len(self.images) > 1:
                if page==1:
                    self.barre.EnableTool(wx.ID_FILE1, False)
                    self.barre.EnableTool(wx.ID_FILE2, False)
                    self.barre.EnableTool(wx.ID_FILE3, True)
                    self.barre.EnableTool(wx.ID_FILE4, True)
                elif page==len(self.images):
                    self.barre.EnableTool(wx.ID_FILE1, True)
                    self.barre.EnableTool(wx.ID_FILE2, True)
                    self.barre.EnableTool(wx.ID_FILE3, False)
                    self.barre.EnableTool(wx.ID_FILE4, False)
                else:
                    self.barre.EnableTool(wx.ID_FILE1, True)
                    self.barre.EnableTool(wx.ID_FILE2, True)
                    self.barre.EnableTool(wx.ID_FILE3, True)
                    self.barre.EnableTool(wx.ID_FILE4, True)
            else:
                self.barre.EnableTool(wx.ID_FILE1, False)
                self.barre.EnableTool(wx.ID_FILE2, False)
                self.barre.EnableTool(wx.ID_FILE3, False)
                self.barre.EnableTool(wx.ID_FILE4, False)
            self.barre.EnableTool(wx.ID_FILE5, True)
            self.barre.EnableTool(wx.ID_FILE6, True)
            self.barre.EnableTool(wx.ID_FILE7, True)
            self.barre.EnableTool(wx.ID_FILE8, True)
            self.barre.EnableTool(wx.ID_FILE9, True)
            self.barre.EnableTool(wx.ID_PRINT, True)
            self.page=page
        else:
            texte = u"Page 0 sur 0"
            self.etiquette.SetLabel(texte)
            self.image=None
            self.affichage.SetImage()
            self.barre.EnableTool(wx.ID_FILE1, False)
            self.barre.EnableTool(wx.ID_FILE2, False)
            self.barre.EnableTool(wx.ID_FILE3, False)
            self.barre.EnableTool(wx.ID_FILE4, False)
            self.barre.EnableTool(wx.ID_FILE5, False)
            self.barre.EnableTool(wx.ID_FILE6, False)
            self.barre.EnableTool(wx.ID_FILE7, False)
            self.barre.EnableTool(wx.ID_FILE8, False)
            self.barre.EnableTool(wx.ID_FILE9, False)
            self.barre.EnableTool(wx.ID_PRINT, False)
            self.page=0
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.loupe=0
        
    def GetImageHandler(self, image):
        eclate=image.split(".")
        term=eclate[len(eclate)-1]
        if term == "jpg":
            return wx.BITMAP_TYPE_JPEG
        elif term == "png":
            return wx.BITMAP_TYPE_PNG
        elif term == "pnm":
            return wx.BITMAP_TYPE_PNM
        elif term == "tif":
            return wx.BITMAP_TYPE_TIF
        else:
            return wx.BITMAP_TYPE_BMP

    def Save(self, event):
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.image = self.bmp
        self.image.SaveFile(self.images[self.page-1], self.GetImageHandler(self.images[self.page-1]))
        self.rotation = 0
        self.bmp=None

    def Rotate(self, event):
        self.rotation += 1
        if self.rotation == 4:
            self.rotation = 0
            self.affichage.SetImage()
            self.affichage.SetImage(self.image)
            self.bmp=None
            self.barre.EnableTool(wx.ID_SAVE, False)
        else:
            img = self.image.ConvertToImage()
            for x in range(self.rotation):
                img = img.Rotate90(True)
            self.bmp = img.ConvertToBitmap()
            self.affichage.SetImage()
            self.affichage.SetImage(self.bmp)
            self.barre.EnableTool(wx.ID_SAVE, True)

    def Print(self, event):
        printData = wx.PrintData()
        printData.SetPaperId(wx.PAPER_A4)
        printData.SetPrintMode(wx.PRINT_MODE_PRINTER)
        pdd = wx.PrintDialogData(printData)
        pdd.SetToPage(1)
        printer = wx.Printer(pdd)
        printout = MyPrintout(self.image)

        printer.Print(GLOBVAR.app, printout, True)
        erreur = printer.GetLastError()
        if erreur == wx.PRINTER_NO_ERROR:
            printData = wx.PrintData(printer.GetPrintDialogData().GetPrintData())
        printout.Destroy()
        if erreur == wx.PRINTER_CANCELLED:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"Impression abandonnée par l'utilisateur",
                                   caption=u"Annulation",
                                   style=wx.OK|wx.ICON_INFORMATION)
            val=dlg.ShowModal()
            dlg.Destroy()
        elif erreur == wx.PRINTER_ERROR:
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=u"Une erreur inattendue est intervenue",
                                   caption=u"Impression impossible",
                                   style=wx.OK|wx.ICON_ERROR)
            val=dlg.ShowModal()
            dlg.Destroy()

    def OnClickFirst(self, event):
        self.loupe = 0
        self.rotation = 0
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.page = 1
        self.image = wx.Bitmap(self.images[self.page - 1], wx.BITMAP_TYPE_ANY)
        self.affichage.SetImage()
        self.affichage.SetImage(self.image)
        self.barre.EnableTool(wx.ID_FILE1, False)
        self.barre.EnableTool(wx.ID_FILE2, False)
        self.barre.EnableTool(wx.ID_FILE3, True)
        self.barre.EnableTool(wx.ID_FILE4, True)
        mess = u"Page %s sur %s"%(self.page, len(self.images))
        self.etiquette.SetLabel(mess)

    def OnClickBack(self, event):
        self.loupe = 0
        self.rotation = 0
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.page = self.page - 1
        self.image = wx.Bitmap(self.images[self.page - 1], wx.BITMAP_TYPE_ANY)
        self.affichage.SetImage()
        self.affichage.SetImage(self.image)
        if self.page == 1:
            self.barre.EnableTool(wx.ID_FILE1, False)
            self.barre.EnableTool(wx.ID_FILE2, False)
            self.barre.EnableTool(wx.ID_FILE3, True)
            self.barre.EnableTool(wx.ID_FILE4, True)
        else:
            self.barre.EnableTool(wx.ID_FILE1, True)
            self.barre.EnableTool(wx.ID_FILE2, True)
            self.barre.EnableTool(wx.ID_FILE3, True)
            self.barre.EnableTool(wx.ID_FILE4, True)
        mess = u"Page %s sur %s"%(self.page, len(self.images))
        self.etiquette.SetLabel(mess)

    def OnClickForward(self, event):
        self.loupe = 0
        self.rotation = 0
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.page = self.page + 1
        self.image = wx.Bitmap(self.images[self.page - 1], wx.BITMAP_TYPE_ANY)
        self.affichage.SetImage()
        self.affichage.SetImage(self.image)
        if self.page == len(self.images):
            self.barre.EnableTool(wx.ID_FILE1, True)
            self.barre.EnableTool(wx.ID_FILE2, True)
            self.barre.EnableTool(wx.ID_FILE3, False)
            self.barre.EnableTool(wx.ID_FILE4, False)
        else:
            self.barre.EnableTool(wx.ID_FILE1, True)
            self.barre.EnableTool(wx.ID_FILE2, True)
            self.barre.EnableTool(wx.ID_FILE3, True)
            self.barre.EnableTool(wx.ID_FILE4, True)
        mess = u"Page %s sur %s"%(self.page, len(self.images))
        self.etiquette.SetLabel(mess)

    def OnClickLast(self, event):
        self.loupe = 0
        self.rotation = 0
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.page = len(self.images)
        self.image = wx.Bitmap(self.images[self.page - 1], wx.BITMAP_TYPE_ANY)
        self.affichage.SetImage()
        self.affichage.SetImage(self.image)
        self.barre.EnableTool(wx.ID_FILE1, True)
        self.barre.EnableTool(wx.ID_FILE2, True)
        self.barre.EnableTool(wx.ID_FILE3, False)
        self.barre.EnableTool(wx.ID_FILE4, False)
        mess = u"Page %s sur %s"%(self.page, len(self.images))
        self.etiquette.SetLabel(mess)

    def Retablir(self, event):
        self.rotation = 0
        self.barre.EnableTool(wx.ID_SAVE, False)
        posX, posY = self.affichage.GetViewStart()
        self.loupe = 0
        self.affichage.SetImage()
        self.affichage.SetImage(self.image)
        self.affichage.Scroll(posX, posY)

    def Ajuster(self, event):
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.rotation = 0
        posX, posY = self.affichage.GetViewStart()
        largFen, hautFen = self.affichage.GetClientSizeTuple()
        img = self.image.ConvertToImage()
        largIm = img.GetWidth()
        diff = abs(largIm - largFen)
        hautIm = img.GetHeight()
        if diff > 0:
            if largFen > largIm:
                ajust = (diff*1. / largFen)
                ratio = 1 + ajust
                self.loupe = int(ajust * 10)
            else:
                ajust = (diff*1. /largIm)
                ratio = 1 - ajust
                self.loupe = int(ajust * 10) * -1
            larg = int(largIm * ratio)
            haut = int(hautIm * ratio)
            img.Rescale(larg, haut)
            bmp = img.ConvertToBitmap()
            self.affichage.SetImage()
            self.affichage.SetImage(bmp)
            self.affichage.Scroll(posX, posY)

    def ZoomerPlus(self, event):
        self.rotation = 0
        posX, posY = self.affichage.GetViewStart()
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.loupe += 1
        ratio = 1 + (self.loupe * 0.10)
        img = self.image.ConvertToImage()
        largIm = img.GetWidth()
        hautIm = img.GetHeight()
        larg = int(largIm * ratio)
        haut = int(hautIm * ratio)
        img.Rescale(larg, haut)
        bmp = img.ConvertToBitmap()
        self.affichage.SetImage()
        self.affichage.SetImage(bmp)
        self.affichage.Scroll(posX, posY)

    def ZoomerMoins(self, event):
        self.rotation = 0
        posX, posY = self.affichage.GetViewStart()
        self.barre.EnableTool(wx.ID_SAVE, False)
        self.loupe -= 1
        ratio = 1 + (self.loupe * 0.10)
        if ratio > 0:
            img = self.image.ConvertToImage()
            largIm = img.GetWidth()
            hautIm = img.GetHeight()
            larg = int(largIm * ratio)
            haut = int(hautIm * ratio)
            img.Rescale(larg, haut)
            bmp = img.ConvertToBitmap()
            self.affichage.SetImage()
            self.affichage.SetImage(bmp)
            self.affichage.Scroll(posX, posY)
        else:
            self.loupe += 1

class Apercu(wx.ScrolledWindow):
    def __init__(self, parent):
        wx.ScrolledWindow.__init__(self, parent=parent, id=-1)
        self.SetBackgroundColour(wx.WHITE)
        self.support=wx.EmptyBitmap(1, 1)
        dc = wx.BufferedDC(None, self.support)
        dc.Clear()
        self.image=None
        self.SetScrollRate(20, 20)
        wx.EVT_PAINT(self, self.OnPaint)

    def OnPaint(self, event):
        dc = wx.BufferedPaintDC(self, self.support, wx.BUFFER_VIRTUAL_AREA)

    def OnImage(self, bmp):
        self.SetVirtualSize((bmp.GetWidth(), bmp.GetHeight()))
        self.dimX = bmp.GetWidth()
        self.dimY = bmp.GetHeight()
        self.support = wx.EmptyBitmap(self.dimX, self.dimY)
        dc = wx.BufferedDC(None, self.support)
        dc.Clear()
        dc.DrawBitmap(bmp, 0, 0)
        self.Refresh()

    def SetImage(self, img = None):
        if img:
            self.OnImage(img)
        else:
            dc = wx.BufferedDC(None, self.support)
            dc.Clear()
            self.SetVirtualSize(self.GetClientSize())
            self.Refresh()

class Calendrier(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent=parent, id=-1, title=u"Choisir une date")
        box=wx.BoxSizer(wx.VERTICAL)
        self.cal = wx.calendar.CalendarCtrl(self, -1, style=wx.calendar.CAL_MONDAY_FIRST)
        box.Add(self.cal, 1, wx.EXPAND|wx.ALL, border=10)
        boutonSizer = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box.Add(boutonSizer, 0, wx.ALIGN_RIGHT|wx.ALL, border=5)
        self.SetSizer(box)
        self.Fit()
        self.SetAutoLayout(True)

    def GetDate(self):
        return self.cal.GetDate()

class Enregistrer(wx.Dialog):
    def __init__(self, parent, mode, fichier=None):
        wx.Dialog.__init__(self, parent = parent, id=-1, title = u"Enregistrer le document")
        self.mode = mode
        box=wx.BoxSizer(wx.VERTICAL)
        self.fichier = fichier
        self.ladate = wx.DateTime().Today()
        eclate = self.ladate.FormatISODate().split("-")
        date_actuelle = "%s/%s/%s"%(eclate[2], eclate[1], eclate[0])
        label1 = wx.StaticText(self, -1, u"Saisir le titre du document", style=wx.ALIGN_CENTRE)
        label2 = wx.StaticText(self, -1, u"Date de classement du document", style=wx.ALIGN_CENTRE)
        self.texte = wx.TextCtrl(self, -1, size=(350, -1))
        self.choix_date = wx.StaticText(self, -1, date_actuelle, style=wx.ALIGN_CENTRE)
        bouton = wx.Button(self, -1, u"Modifier")
        self.Bind(wx.EVT_BUTTON, self.Modifier, bouton)
        self.chkPdf = wx.CheckBox(self, -1, u"Numériser en PDF")
        self.Bind(wx.EVT_CHECKBOX, self.ChoisirPDF, self.chkPdf)
        self.radio1 = wx.RadioButton(self, -1, u"Au format A4")
        self.radio2 = wx.RadioButton(self, -1, u"En pleine résolution")
        visu = wx.Button(self, -1, u"Vérifier le document PDF")
        self.Bind(wx.EVT_BUTTON, self.Visualiser, visu)
        self.radio1.Enable(False)
        self.radio2.Enable(False)
        if mode != 3:
            self.chkPdf.Show(False)
            self.radio1.Show(False)
            self.radio2.Show(False)
        if mode == 1:
            visu.Enable(True)
        if mode == 2:
            visu.SetLabel(u"Vérifier le document Open-Office")
            visu.Enable(True)
        if mode == 0 or mode == 3:
            visu.Enable(False)
        box1 = wx.BoxSizer(wx.HORIZONTAL)
        box1.Add(label1, 1 , wx.ALL, border=5)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        box2.Add(self.texte, 1, wx.ALL, border=5)
        box3 = wx.BoxSizer(wx.HORIZONTAL)
        box3.Add(label2, 1, wx.ALL, border=5)
        box4 = wx.BoxSizer(wx.HORIZONTAL)
        box4.Add(self.choix_date, 1, wx.ALIGN_CENTRE|wx.ALL, border=5)
        box4.Add(bouton, 1, wx.ALL, border=5)
        box.Add(box1, 0, wx.EXPAND)
        box.Add(box2, 0, wx.EXPAND)
        box.Add(box3, 0, wx.EXPAND)
        box.Add(box4, 0, wx.EXPAND)
        box.Add(self.chkPdf, 0, wx.EXPAND|wx.ALL, border=5)
        box.Add(self.radio1, 0, wx.EXPAND|wx.ALL, border=5)
        box.Add(self.radio2, 0, wx.EXPAND|wx.ALL, border=5)
        box.Add(visu, 0, wx.EXPAND|wx.ALL, border=5)
        sizerButtons = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box.Add(sizerButtons, 0, wx.ALIGN_RIGHT|wx.TOP, border=5)
        self.SetSizer(box)
        box.Fit(self)
        self.SetAutoLayout(True)

    def Visualiser(self, event):
        if WIN:
            os.startfile(self.fichier)
        else:
            if self.mode == 2:
                commande = "ooffice \"%s\""%self.fichier
            else:
                commande = GLOBVAR.visupdf + " \"%s\""%self.fichier
            os.system(commande)

    def ChoisirPDF(self, event):
        if self.chkPdf.IsChecked():
            self.radio1.Enable(True)
            self.radio2.Enable(True)
        else:
            self.radio1.Enable(False)
            self.radio2.Enable(False)
        self.Layout()

    def Modifier(self, event):
        dlg = Calendrier(self)
        rep=dlg.ShowModal()
        nouvDate = dlg.GetDate()
        dlg.Destroy()
        if rep == wx.ID_OK:
            if nouvDate <= wx.DateTime().Today():
                self.ladate = nouvDate
                eclate = nouvDate.FormatISODate().split("-")
                nouvDate = "%s/%s/%s"%(eclate[2], eclate[1], eclate[0])
                self.choix_date.SetLabel(nouvDate)
                self.Layout()

    def GetTitre(self):
        return self.texte.GetValue()

    def GetDate(self):
        return self.ladate

    def GetPDF(self):
        return (self.chkPdf.IsChecked(), self.radio1.GetValue())

class AjoutFichier(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent, id=-1)

        self.serieEnCours = False
        self.liste = []
        self.page = 0
        self.choixPDF = False
        self.txt = False

        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        self.btCommencer = wx.Button(self, -1, u'Commencer')
        self.Bind(wx.EVT_BUTTON, self.Commencer, self.btCommencer)
        box2.Add(self.btCommencer, 1, wx.ALL, border= 5)

        self.btEnregistrer = wx.Button(self, -1, u'Enregistrer')
        self.Bind(wx.EVT_BUTTON, self.Enreg, self.btEnregistrer)
        box2.Add(self.btEnregistrer, 1, wx.ALL, border= 5)

        self.btAnnuler = wx.Button(self, -1, u'Annuler')
        self.Bind(wx.EVT_BUTTON, self.Annuler, self.btAnnuler)
        box2.Add(self.btAnnuler, 1, wx.ALL, border= 5)

        box1.Add(box2, 0, wx.EXPAND)
        self.panneau = Affichage(self)
        box1.Add(self.panneau, 1, wx.EXPAND)
        self.SetSizer(box1)
        self.SetAutoLayout(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)

    def Annuler(self, event):
        self.panneau.SetDocument()
        self.btCommencer.Enable(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)

    def Enreg(self, event):
        self.Enregistrer()

    def Enregistrer(self, pdf = False, path = None, ooo = "non"):
        if pdf:
            mode = 1
        elif ooo != "non":
            mode = 2
        else:
            mode = 0
        dlgTxt = Enregistrer(GLOBVAR.app, mode, path)
        val = dlgTxt.ShowModal()
        resultat1 = dlgTxt.GetTitre()
        resultat2 = dlgTxt.GetDate()
        dlgTxt.Destroy()
        if val == wx.ID_OK:
            if resultat1 != "":
                leTitre = "''".join(resultat1.split("'"))
                leTitre=eval('u"%s"'%leTitre)
                laDate = resultat2.FormatISODate()
                lAnnee = laDate.split("-")[0]
                leMois = laDate.split("-")[1]
                if pdf or (ooo != "non"):
                    nbPages = "1"
                else:
                    nbPages = str(len(self.liste))
                c = GLOBVAR.base.cursor()
                req = "INSERT INTO documents(classeur, dossier, chemise, date, titre, nbpages, annee, mois) "
                req = req + "VALUES(0, 0, 0,'%s', '%s', %s, '%s', '%s')"%(laDate, leTitre, nbPages, lAnnee, leMois)
                c.execute(req)
                res = c.execute("SELECT MAX(enreg) FROM documents")
                numero = str(res.fetchone()[0])
                i = 0
                chem = os.path.join(GLOBVAR.docdir, lAnnee)
                if os.path.isdir(chem)== False:
                    os.mkdir(chem)
                chem = os.path.join(chem, leMois)
                if os.path.isdir(chem) == False :
                    os.mkdir(chem)
                if pdf:
                    fichier = numero + "-1.pdf"
                    pathComplet = os.path.join(chem, fichier)
                    shutil.copyfile(path, pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document PDF a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                elif ooo != "non":
                    fichier = numero + "-1." + ooo
                    pathComplet = os.path.join(chem, fichier)
                    shutil.copyfile(path, pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document Open Office a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                else:
                    i = 0
                    for x in self.liste:
                        i += 1
                        eclate = x.split(".")
                        term = eclate[len(eclate)-1]
                        fichier = numero + "-" + str(i) + "." + term.lower()
                        pathComplet = os.path.join(chem, fichier)
                        shutil.copyfile(x, pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document complet a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                self.panneau.SetDocument()
                self.btCommencer.Enable(True)
                self.btEnregistrer.Enable(False)
                self.btAnnuler.Enable(False)

    def Commencer(self, event):
        self.liste = []
        boucle = True
        ficImg = False
        mess = u"Choisissez un fichier image, un PDF ou un document Open-Office"
        while boucle :
            dlg = wx.FileDialog(parent=GLOBVAR.app,
                                message=mess)
            val = dlg.ShowModal()
            fichier = dlg.GetPath()
            dlg.Destroy()
            if val == wx.ID_OK :
                if fichier != "":
                    eclate = fichier.split(".")
                    terminaison = eclate[len(eclate)-1]
                    if terminaison.upper() in GLOBVAR.listedoc :
                        if terminaison.upper() == "PDF" :
                            if ficImg:
                                dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                                       message=u"Le format n'est pas compatible avec celui des pages précédentes",
                                                       caption = u"Impossible",
                                                       style= wx.OK|wx.ICON_ERROR)
                                val = dlg.run()
                                dlg.destroy()
                            else:
                                self.Enregistrer(True, fichier, "non")
                                return
                        else:
                            ficImg = True
                            mess = u"Choisissez un fichier image"
                            self.liste.append(fichier)
                            self.panneau.SetDocument(self.liste, page = len(self.liste))
                            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                                   message=u"Voulez-vous ajouter une page à votre document ?",
                                                   caption = u"Ajout manuel",
                                                   style = wx.YES_NO|wx.ICON_QUESTION)
                            val = dlg.ShowModal()
                            dlg.Destroy()
                            if val == wx.ID_NO :
                                boucle = False
                    elif terminaison.upper() in GLOBVAR.listeoo :
                        if ficImg:
                            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                                   message=u"Le format n'est pas compatible avec celui des pages précédentes",
                                                   caption = u"Impossible",
                                                   style= wx.OK|wx.ICON_ERROR)
                            val = dlg.ShowModal()
                            dlg.Destroy()
                        else:
                            self.Enregistrer(False, fichier, terminaison)
                            return
                    else:
                        if ficImg:
                            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                                   message=u"Le format n'est pas compatible avec celui des pages précédentes",
                                                   caption = u"Impossible",
                                                   style= wx.OK|wx.ICON_ERROR)
                            val = dlg.ShowModal()
                            dlg.Destroy()
                        else:
                            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                                   message=u"Format de document invalide",
                                                   caption = u"Impossible",
                                                   style= wx.OK|wx.ICON_ERROR)
                            val = dlg.ShowModal()
                            dlg.Destroy()
                            boucle = False
                else:
                    boucle = False
            else:
                boucle = False
        if len(self.liste) != 0:
            self.btCommencer.Enable(False)
            self.btEnregistrer.Enable(True)
            self.btAnnuler.Enable(True)
            
class Note(wx.Dialog):
    def __init__(self, titre):
        l, h = wx.ScreenDC().GetSizeTuple()
        largeur = h/2
        hauteur = l/2
        taille=wx.Size(largeur, hauteur)
        wx.Dialog.__init__(self, parent=GLOBVAR.app, title=titre, size=taille)
        self.pdf = False
        box=wx.BoxSizer(wx.VERTICAL)
        labelTitre = wx.StaticText(self, -1, u"Saisissez votre texte puis validez", style=wx.ALIGN_CENTRE)
        box.Add(labelTitre, 0, wx.EXPAND|wx.ALL, border=4)
        panneau = wx.ScrolledWindow(self, -1)
        self.ctrlTexte = wx.TextCtrl(panneau, -1, style=wx.TE_MULTILINE)
        sizer=wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.ctrlTexte, 1, wx.EXPAND)
        panneau.SetSizer(sizer)
        panneau.SetAutoLayout(True)
        box.Add(panneau, 1, wx.EXPAND)
        self.chkPdf = wx.CheckBox(self, -1, u"Sauvegarder en PDF")
        self.Bind(wx.EVT_CHECKBOX, self.ChoisirPDF, self.chkPdf)
        box.Add(self.chkPdf, 0, wx.ALIGN_LEFT|wx.ALL, border=4)
        sizerButtons = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box.Add(sizerButtons, 0, wx.ALIGN_RIGHT|wx.ALL, border=5)
        self.SetSizer(box)
        self.SetAutoLayout(True)

    def GetText(self):
        return self.ctrlTexte.GetValue()

    def ChoisirPDF(self, event):
        if self.chkPdf.IsChecked():
            self.pdf = True
        else:
            self.pdf = False

    def FormatPDF(self):
        return self.pdf

class AffichTexte(wx.Panel): 
    def __init__(self, parent): 
        wx.Panel.__init__(self, parent = parent, id=-1)  
        self.fichier = None
        self.box = wx.BoxSizer(wx.VERTICAL)  
        self.label = wx.StaticText(self, -1, u"Note manuscrite", style= wx.ALIGN_CENTRE)
        self.box.Add(self.label, 0, wx.EXPAND)
        self.barre = wx.ToolBar(self, -1)
        self.barre.SetToolBitmapSize((32, 32))
        self.imprim_bouton = self.barre.AddSimpleTool(wx.ID_PRINT,
                                                      wx.Bitmap(os.path.join(GLOBVAR.themedir, "imprimer.png"), wx.BITMAP_TYPE_PNG),
                                                      shortHelpString = u"Imprimer le document via un PDF")
        self.barre.Realize()
        self.box.Add(self.barre, flag=wx.EXPAND|wx.ALL, border=4)
        self.ecran = wx.ScrolledWindow(self, -1)
        boxEcran = wx.BoxSizer(wx.VERTICAL)
        textStyle = wx.TE_MULTILINE|wx.TE_READONLY
        self.affichage = wx.TextCtrl(self.ecran, -1, style=textStyle)
        boxEcran.Add(self.affichage, 1, wx.EXPAND)
        self.ecran.SetSizer(boxEcran)
        self.ecran.SetAutoLayout(True)
        self.box.Add(self.ecran, 1, wx.EXPAND|wx.ALL, border=4)
        self.SetSizer(self.box)
        self.SetAutoLayout(True)
        
        self.barre.EnableTool(wx.ID_PRINT, False)
        
        wx.EVT_TOOL(self, wx.ID_PRINT, self.Imprimer)

    def SetDocument(self, fichier=None):
        if fichier:
            self.fichier = fichier
            self.affichage.LoadFile(self.fichier)  
            self.barre.EnableTool(wx.ID_PRINT, True)
        else:
            self.affichage.Clear()
            self.barre.EnableTool(wx.ID_PRINT, False)

    def Imprimer(self, event):
        fic = open(self.fichier, "r")
        texte = fic.read()
        fic.close()
        fichier = os.path.join(GLOBVAR.tempdir, "temp.pdf")
        c = canvas.Canvas(fichier, pagesize = A4)
        w, h = A4
        c.setFont("Helvetica-Bold", 14)
        chaine = u"Note manuscrite"
        posX = (w - c.stringWidth(chaine, "Helvetica-Bold", 14))/2
        posY = h-(2*cm)
        c.drawString(posX, posY, chaine)
        c.setFont("Helvetica", 12)
        phrases = texte.split("\n")
        posX = 2*cm
        posY = h-(3*cm)
        maxLigne = w-(4*cm)
        hauteur = 0.5*cm
        nbrelignes = int((h-(5*cm)) / hauteur)
        ligne = 0
        page = 0
        for x in phrases :
            ligne = ligne + 1
            if ligne > nbrelignes :
                ligne = 1
                c.showPage()
                posX = 2*cm
                posY = h-(2*cm)
                nbrelignes = int((h-(4*cm)) / hauteur)
                c.setFont("Helvetica", 12)
                page = page + 1
            mots = x.split()
            chaine = ""
            for y in mots:
                temporaire = chaine + y + " "
                if c.stringWidth(temporaire, "Helvetica", 12) > maxLigne :
                    c.drawString(posX, posY, chaine)
                    posY = posY - hauteur
                    chaine = y + " "
                    ligne = ligne + 1
                else:
                    chaine = temporaire
            c.drawString(posX, posY, chaine)
            posY = posY - hauteur
        c.showPage()
        c.save()
        if WIN:
            os.startfile(fichier)
        else:
            os.system(GLOBVAR.visupdf + " '" + fichier + "'")
        
class AffichageTextesDialog(wx.Dialog):
    def __init__(self, titre, fichier):
        l, h = wx.ScreenDC().GetSizeTuple()
        largeur = h/2
        hauteur = l/2
        taille = wx.Size(largeur, hauteur)
        wx.Dialog.__init__(self, parent=GLOBVAR.app, title = titre, size=taille)
        self.fichier = fichier
        box=wx.BoxSizer(wx.VERTICAL)
        affich = AffichTexte(self)
        box.Add(affich, 1, wx.EXPAND)
        affich.SetDocument(self.fichier)
        sizerbouton = self.CreateButtonSizer(wx.OK)
        box.Add(sizerbouton, flag=wx.ALIGN_RIGHT|wx.ALL, border=10)
        self.SetSizer(box)
        self.SetAutoLayout(True)
        
class AjoutNote(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent, id=-1)

        self.choixPDF = False
        self.liste=[]

        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        self.btCommencer = wx.Button(self, -1, u'Commencer')
        self.Bind(wx.EVT_BUTTON, self.Commencer, self.btCommencer)
        box2.Add(self.btCommencer, 1, wx.ALL, border= 5)

        self.btEnregistrer = wx.Button(self, -1, u'Enregistrer')
        self.Bind(wx.EVT_BUTTON, self.Enreg, self.btEnregistrer)
        box2.Add(self.btEnregistrer, 1, wx.ALL, border= 5)

        self.btAnnuler = wx.Button(self, -1, u'Annuler')
        self.Bind(wx.EVT_BUTTON, self.Annuler, self.btAnnuler)
        box2.Add(self.btAnnuler, 1, wx.ALL, border= 5)

        box1.Add(box2, 0, wx.EXPAND)
        self.panneau = AffichTexte(self)
        box1.Add(self.panneau, 1, wx.EXPAND)
        self.SetSizer(box1)
        self.SetAutoLayout(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)

    def Annuler(self, event):
        self.panneau.SetDocument()
        self.btCommencer.Enable(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)
        
    def Commencer(self, event):
        dlg = Note(u"Note manuscrite")
        rep = dlg.ShowModal()
        texte = dlg.GetText()
        pdf = dlg.FormatPDF()
        dlg.Destroy()
        if rep == wx.ID_OK and texte != "":
            if pdf:
                fichier = os.path.join(GLOBVAR.tempdir, "temp.pdf")
                c = canvas.Canvas(fichier, pagesize = A4)
                w, h = A4
                c.setFont("Helvetica-Bold", 14)
                chaine = u"Note manuscrite"
                posX = (w - c.stringWidth(chaine, "Helvetica-Bold", 14))/2
                posY = h-(2*cm)
                c.drawString(posX, posY, chaine)
                c.setFont("Helvetica", 12)
                if "\r" in chaine:
                    phrases = texte.split("\r\n")
                else:   
                    phrases = texte.split("\n")
                posX = 2*cm
                posY = h-(3*cm)
                maxLigne = w-(4*cm)
                hauteur = 0.5*cm
                nbrelignes = int((h-(5*cm)) / hauteur)
                ligne = 0
                page = 0
                for x in phrases :
                    ligne = ligne + 1
                    if ligne > nbrelignes :
                        ligne = 1
                        c.showPage()
                        posX = 2*cm
                        posY = h-(2*cm)
                        nbrelignes = int((h-(4*cm)) / hauteur)
                        c.setFont("Helvetica", 12)
                        page = page + 1
                    mots = x.split()
                    chaine = ""
                    for y in mots:
                        temporaire = chaine + y + " "
                        if c.stringWidth(temporaire, "Helvetica", 12) > maxLigne :
                            c.drawString(posX, posY, chaine)
                            posY = posY - hauteur
                            chaine = y + " "
                            ligne = ligne + 1
                        else:
                            chaine = temporaire
                    c.drawString(posX, posY, chaine)
                    posY = posY - hauteur
                c.showPage()
                c.save()
                self.Enregistrer(True, fichier)
            else:
                texte = texte.encode('utf-8')
                self.liste = []
                fichier = os.path.join(GLOBVAR.tempdir, "temp.txt")
                fic = open(fichier, "w")
                fic.write(texte)
                fic.close()
                self.liste.append(fichier)
                self.panneau.SetDocument(fichier)
                self.btEnregistrer.Enable(True)
                self.btAnnuler.Enable(True)
                self.btCommencer.Enable(False)
            
    def Enreg(self, event):
        self.Enregistrer()
            
    def Enregistrer(self, pdf = False, path = None):
        mode = 0
        dlgTxt = Enregistrer(GLOBVAR.app, mode, path)
        val = dlgTxt.ShowModal()
        resultat1 = dlgTxt.GetTitre()
        resultat2 = dlgTxt.GetDate()
        dlgTxt.Destroy()
        if val == wx.ID_OK:
            if resultat1 != "":
                leTitre = "''".join(resultat1.split("'"))
                leTitre=eval('u"%s"'%leTitre)
                laDate = resultat2.FormatISODate()
                lAnnee = laDate.split("-")[0]
                leMois = laDate.split("-")[1]
                nbPages = "1"
                c = GLOBVAR.base.cursor()
                req = "INSERT INTO documents(classeur, dossier, chemise, date, titre, nbpages, annee, mois) "
                req = req + "VALUES(0, 0, 0,'%s', '%s', %s, '%s', '%s')"%(laDate, leTitre, nbPages, lAnnee, leMois)
                c.execute(req)
                res = c.execute("SELECT MAX(enreg) FROM documents")
                numero = str(res.fetchone()[0])
                i = 0
                chem = os.path.join(GLOBVAR.docdir, lAnnee)
                if os.path.isdir(chem)== False:
                    os.mkdir(chem)
                chem = os.path.join(chem, leMois)
                if os.path.isdir(chem) == False :
                    os.mkdir(chem)
                if pdf:
                    fichier = numero + "-1.pdf"
                    pathComplet = os.path.join(chem, fichier)
                    shutil.copyfile(path, pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document PDF a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                else:
                    fichier = numero + "-1.txt"
                    pathComplet = os.path.join(chem, fichier)
                    shutil.copyfile(self.liste[0], pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document texte a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                self.panneau.SetDocument()
                self.btCommencer.Enable(True)
                self.btEnregistrer.Enable(False)
                self.btAnnuler.Enable(False)
                
class ChoixChemise(wx.Panel):
    def __init__(self, parent):
        (largeur, hauteur) = parent.GetClientSizeTuple()
        taille = wx.Size(largeur, hauteur)
        wx.Panel.__init__(self, parent=parent, id=-1, size=taille)
        self.origine = parent
        self.valeur = []
        self.fenetre = wx.ScrolledWindow(self, -1)
        fenSizer = wx.BoxSizer(wx.VERTICAL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.listeImage = wx.ImageList(24, 24)
        self.imRoot = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "racine.png"), wx.BITMAP_TYPE_PNG))
        self.imClasseur = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "classeur.png"), wx.BITMAP_TYPE_PNG))
        self.imDossier = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "dossier.png"), wx.BITMAP_TYPE_PNG))
        self.imChemise = self.listeImage.Add(wx.Bitmap(os.path.join(GLOBVAR.themedir, "chemise.png"), wx.BITMAP_TYPE_PNG))
        self.arbre = wx.TreeCtrl(self.fenetre, -1)
        fenSizer.Add(self.arbre, 1, wx.EXPAND)
        self.fenetre.SetSizer(fenSizer)
        self.fenetre.SetAutoLayout(True)
        self.arbre.SetImageList(self.listeImage)
        self.c = GLOBVAR.base.cursor()
        self.Remplir()
        sizer.Add(sizer2, flag=wx.EXPAND|wx.ALL, border=5)
        sizer.Add(self.fenetre, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.SetAutoLayout(True)

        self.fenetre.SetScrollRate(20, 20)

        self.arbre.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelect)

    def Remplir(self):
        self.root = self.arbre.AddRoot(u"Plan de classement des documents", self.imRoot)
        self.itemChoisi = self.root

        self.c.execute("SELECT COUNT(*) FROM classeurs")
        if (self.c.fetchall()[0][0] > 0):
            self.c.execute("SELECT classeur, libelle FROM classeurs ORDER BY classeur")
            listeClasseurs = self.c.fetchall()

            for x in listeClasseurs:
                leClasseur = x[0]
                leLibelle = x[1]
                myData = wx.TreeItemData([leClasseur, leLibelle])
                child1 = self.arbre.AppendItem(self.root, leLibelle, self.imClasseur, data=myData)
                req = "SELECT dossier, libelle FROM dossiers WHERE classeur = %s ORDER BY dossier"%leClasseur
                self.c.execute(req)
                listeDossiers = self.c.fetchall()

                for y in listeDossiers:
                    leDossier = y[0]
                    leLibelle = y[1]
                    myData = wx.TreeItemData([leClasseur, leDossier, leLibelle])
                    child2 = self.arbre.AppendItem(child1, leLibelle, self.imDossier, data=myData)
                    req = "SELECT chemise, libelle FROM chemises where classeur = %s AND dossier = %s ORDER BY chemise"%(leClasseur, leDossier)
                    self.c.execute(req)
                    listeChemises = self.c.fetchall()

                    for z in listeChemises:
                        laChemise = z[0]
                        leLibelle = z[1]
                        myData = wx.TreeItemData([leClasseur, leDossier, laChemise, leLibelle])
                        child3 = self.arbre.AppendItem(child2, leLibelle, self.imChemise, data=myData)
            self.arbre.Expand(self.root)

    def OnSelect(self, event):
        item = event.GetItem()
        if (self.arbre.GetItemImage(item) == self.imChemise) :
            self.valeur=[self.arbre.GetPyData(item)[0], self.arbre.GetPyData(item)[1], self.arbre.GetPyData(item)[2]]
            self.c.execute("SELECT libelle FROM classeurs WHERE classeur = %s"%self.valeur[0])
            leclasseur = self.c.fetchone()[0] 
            self.c.execute("SELECT libelle FROM dossiers WHERE dossier = %s"%self.valeur[1])
            ledossier = self.c.fetchone()[0] 
            self.origine.SetText(u"Chemise choisie : %s - %s - %s"%(leclasseur, ledossier, self.arbre.GetPyData(item)[3]))
        else :
            self.valeur = []
            self.origine.SetText(u"Choisissez une chemise de classement")
        self.Layout()
        
    def GetValue(self):
        return self.valeur    
                
class ClassementDialog(wx.Dialog):
    def __init__(self, parent, titre):
        l, h = wx.ScreenDC().GetSizeTuple()
        larg = (l * 2) / 3
        haut = (h * 2) / 3
        taille = wx.Size(larg, haut)
        wx.Dialog.__init__(self, parent=parent, title=titre, size=taille)
        self.curseur = GLOBVAR.base.cursor()
        box = wx.BoxSizer(wx.VERTICAL)
        self.label = wx.StaticText(self, -1, u"Choisissez une chemise de classement", style=wx.ALIGN_CENTRE)
        box.Add(self.label, 0, wx.EXPAND|wx.ALL, border=4)
        self.affich = ChoixChemise(self)
        box.Add(self.affich, 1, wx.EXPAND)
        sizerButtons = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box.Add(sizerButtons, 0, wx.ALIGN_RIGHT|wx.ALL, border=5)
        self.SetSizer(box)
        self.SetAutoLayout(True)

    def SetText(self, chaine):
        self.label.SetLabel(chaine)
        self.Layout()
        
    def GetValue(self):
        return self.affich.GetValue()
                
class AClasser(wx.Panel):
    def __init__(self, parent, origine):
        wx.Panel.__init__(self, parent=parent, id=-1, style=wx.BORDER_SUNKEN)
        self.origine = origine
        self.item = None
        self.resultat=[]
        self.selection=[]
        boxPrinc = wx.BoxSizer(wx.VERTICAL)
        self.c = GLOBVAR.base.cursor()
        etiquette = wx.StaticText(self,
                                  id= -1,
                                  label = u"Vous accéderez aux menus en cliquant\n les items avec le bouton droit de la souris",
                                  style = wx.ALIGN_CENTRE)
        boxPrinc.Add(etiquette, flag=wx.EXPAND|wx.ALL, border=5)
        self.grille = wx.ScrolledWindow(self, -1)
        boxGrid = wx.BoxSizer(wx.VERTICAL)
        self.listResultat = wx.ListCtrl(self.grille, -1, style=wx.LC_REPORT)
        self.listResultat.InsertColumn(0, u"Date")
        self.listResultat.InsertColumn(1, u"Titre du document")
        boxGrid.Add(self.listResultat, 1, wx.EXPAND)
        self.grille.SetSizer(boxGrid)
        self.grille.SetAutoLayout(True)
        boxPrinc.Add(self.grille, 1, flag=wx.EXPAND|wx.ALL, border=5)
        self.SetSizer(boxPrinc)
        boxPrinc.Fit(self)
        self.SetAutoLayout(True)
        self.popMenuDoc= wx.Menu()
        self.popMenuDoc.Append(wx.ID_FILE1, u"Visualiser le document")
        self.popMenuDoc.Append(wx.ID_FILE2, u"Classer le(s) document(s)")
        self.popMenuDoc.Append(wx.ID_FILE3, u"Supprimer le(s) document(s)")
        self.Rechercher()
        
        wx.EVT_SIZE(self, self.OnSize)
        self.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, self.OnRightClick, self.listResultat)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnSelect, self.listResultat)
        self.Bind(wx.EVT_LIST_ITEM_DESELECTED, self.OnDeSelect, self.listResultat)
        self.Bind(wx.EVT_MENU, self.Visualiser, id = wx.ID_FILE1)
        self.Bind(wx.EVT_MENU, self.Classer, id = wx.ID_FILE2)
        self.Bind(wx.EVT_MENU, self.Supprimer, id = wx.ID_FILE3)

    def OnRightClick(self, event):
        item, flag = self.listResultat.HitTest(event.GetPosition())
        if item > -1:
            pt = event.GetPoint()
            self.listResultat.PopupMenu(self.popMenuDoc, pt)

    def OnSelect(self, event):
        self.origine.SetImage()
        self.selection = []
        for x in range(self.listResultat.GetItemCount()):
            if self.listResultat.GetItemState(x, wx.LIST_STATE_SELECTED):
                self.selection.append(x)
        nombre = self.listResultat.GetSelectedItemCount()
        if nombre > 1:
            self.popMenuDoc.SetLabel(wx.ID_FILE2, u"Classer les %s documents"%nombre)
            self.popMenuDoc.SetLabel(wx.ID_FILE3, u"Supprimer les %s documents"%nombre)
            self.popMenuDoc.Enable(wx.ID_FILE1, False)
            self.popMenuDoc.Enable(wx.ID_FILE2, True)
            self.popMenuDoc.Enable(wx.ID_FILE3, True)
        else:
            self.popMenuDoc.SetLabel(wx.ID_FILE2, u"Classer le document")
            self.popMenuDoc.SetLabel(wx.ID_FILE3, u"Supprimer le documents")
            self.popMenuDoc.Enable(wx.ID_FILE1, True)
            self.popMenuDoc.Enable(wx.ID_FILE2, True)
            self.popMenuDoc.Enable(wx.ID_FILE3, True)
        
    def OnDeSelect(self, event):
        self.origine.SetImage()
        self.selection = []
        for x in range(self.listResultat.GetItemCount()):
            if self.listResultat.GetItemState(x, wx.LIST_STATE_SELECTED):
                self.selection.append(x)
        nombre = self.listResultat.GetSelectedItemCount()
        if nombre == 0:
            self.popMenuDoc.Enable(wx.ID_FILE1, False)
            self.popMenuDoc.Enable(wx.ID_FILE2, False)
            self.popMenuDoc.Enable(wx.ID_FILE3, False)
        elif nombre == 1:
            self.popMenuDoc.SetLabel(wx.ID_FILE2, u"Classer le document")
            self.popMenuDoc.SetLabel(wx.ID_FILE3, u"Supprimer le documents")
            self.popMenuDoc.Enable(wx.ID_FILE1, True)
            self.popMenuDoc.Enable(wx.ID_FILE2, True)
            self.popMenuDoc.Enable(wx.ID_FILE3, True)
        else:
            self.popMenuDoc.SetLabel(wx.ID_FILE2, u"Classer les %s documents"%nombre)
            self.popMenuDoc.SetLabel(wx.ID_FILE3, u"Supprimer les %s documents"%nombre)
            self.popMenuDoc.Enable(wx.ID_FILE1, False)
            self.popMenuDoc.Enable(wx.ID_FILE2, True)
            self.popMenuDoc.Enable(wx.ID_FILE3, True)

    def OnSize(self, event):
        self.Layout()
        larg, haut = self.grille.GetClientSizeTuple()
        larg = larg
        if self.listResultat.GetItemCount() == 0:
            self.listResultat.SetColumnWidth(0, int(larg/4))
            self.listResultat.SetColumnWidth(1, int((larg/4)*3))
        else:
            self.listResultat.SetColumnWidth(0, wx.LIST_AUTOSIZE)
            self.listResultat.SetColumnWidth(1, wx.LIST_AUTOSIZE)
            larg1 = self.listResultat.GetColumnWidth(0)
            larg2 = self.listResultat.GetColumnWidth(1)
            if (larg1 + larg2) < larg:
                self.listResultat.SetColumnWidth(1, larg-larg1)

    def Classer(self, event):
        dlg = ClassementDialog(GLOBVAR.app, u"Classement de document(s)")
        val = dlg.ShowModal()
        resu = dlg.GetValue()
        dlg.Destroy()
        if val == wx.ID_OK and len(resu) == 3:
            for ligne in self.selection:
                enreg = self.resultat[ligne][2]
                req = "UPDATE documents SET classeur = %s, dossier = %s, chemise = %s WHERE enreg = %s"%(resu[0], resu[1], resu[2], enreg)
                self.c.execute(req)
            if len(self.selection) > 1:
                mess = u"Les %s documents ont été classés"%len(self.selection)
            else:
                mess = u"Le document a été classé"
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   caption = u"Opération terminée",
                                   message = mess,
                                   style = wx.OK|wx.ICON_INFORMATION)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.selection = []
            self.listResultat.DeleteAllItems()
            self.origine.SetImage()
            self.Rechercher()
        
    def Rechercher(self):
        self.popMenuDoc.Enable(wx.ID_FILE1, False)
        self.popMenuDoc.Enable(wx.ID_FILE2, False)
        self.popMenuDoc.Enable(wx.ID_FILE3, False)
        req = "SELECT annee, mois, enreg, date, titre, nbpages FROM documents WHERE classeur = 0"
        self.c.execute(req)
        liste = self.c.fetchall()
        if len(liste) > 0:
            self.resultat = liste
            for x in range(len(self.resultat)):
                eclate = self.resultat[x][3].split("-")
                date = eclate[2] + "/" + eclate[1] + "/" + eclate[0]
                index = self.listResultat.InsertStringItem(x, date)
                self.listResultat.SetStringItem(index, 0, date)
                self.listResultat.SetStringItem(index, 1, self.resultat[x][4])
            self.SendSizeEvent()
            return
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               caption = u"Terminé",
                               message = u"Il ne reste plus aucun document à classer",
                               style = wx.OK|wx.ICON_INFORMATION)
        val = dlg.ShowModal()
        dlg.Destroy()

    def Visualiser(self, event):
        titre = self.resultat[self.selection[0]][4]
        enreg = self.resultat[self.selection[0]][2]
        mois = self.resultat[self.selection[0]][1]
        annee = self.resultat[self.selection[0]][0]
        pages = self.resultat[self.selection[0]][5]
        racine = os.path.join(GLOBVAR.docdir, annee, mois)
        if pages == 1:
            fichier = "%s-1"%enreg
            if os.path.isfile(os.path.join(racine,  fichier + ".txt")):
                fic = os.path.join(racine,  fichier + ".txt")
                dlg = AffichageTextesDialog(titre, fic)
                val = dlg.ShowModal()
                dlg.Destroy()
            elif IsItImage(os.path.join(racine,  fichier)):
                liste = []
                liste.append(IsItImage(os.path.join(racine,  fichier)))
                self.origine.SetImage(liste)
            elif IsItOoffice(os.path.join(racine,  fichier)):
                chemin = IsItOoffice(os.path.join(racine,  fichier))
                if WIN:
                    os.startfile(chemin)
                else:
                    commande= 'ooffice "%s"'%chemin
                    os.system(commande)
            else:
                fichier = fichier + ".pdf"
                if WIN:
                    os.startfile(os.path.join(racine,  fichier))
                else:
                    commande= '%s "%s"'%(GLOBVAR.visupdf, os.path.join(racine,  fichier))
                    os.system(commande)
        else:
            maListe = []
            fic = str(enreg) + "-"
            for x in range(pages):
                fichier = IsItImage(os.path.join(racine,  fic + str(x+1)))
                maListe.append(fichier)
            self.origine.SetImage(maListe)

    def Supprimer(self, event):
        nombre = len(self.selection)
        if nombre > 1:
            mess1 = u"Voulez-vous vraiment supprimer ces %s documents ?"%nombre
            mess2 = u"Les %s documents ont été supprimés"%nombre
        else:
            mess1 = u"Voulez-vous vraiment supprimer ce document ?"
            mess2 = u"Le document a été supprimé"
        dlg = wx.MessageDialog(parent=GLOBVAR.app,
                               message=mess1,
                               caption=u"Suppression",
                               style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES:
            for ligne in self.selection:
                titre = self.resultat[ligne][4]
                enreg = self.resultat[ligne][2]
                mois = self.resultat[ligne][1]
                annee = self.resultat[ligne][0]
                pages = self.resultat[ligne][5]
                racine = os.path.join(GLOBVAR.docdir, annee, mois)
                fic = str(enreg) +"-"
                ooo = False
                for term in GLOBVAR.listeoo:
                    fin = term.lower()
                    if os.path.isfile(os.path.join(racine, fic + "1." + fin)):
                        fichier = os.path.join(racine, fic + "1." + fin)
                        ooo = True
                        terminaison = fin
                if ooo:
                    os.remove(os.path.join(racine, fic + "1." + terminaison))
                elif os.path.isfile(os.path.join(racine, fic + "1.pdf")):
                    os.remove(os.path.join(racine, fic + "1.pdf"))
                elif os.path.isfile(os.path.join(racine, fic + "1.txt")):
                    os.remove(os.path.join(racine, fic + "1.txt"))
                else:
                    for x in range(pages):
                        doc = os.path.join(racine, fic + str(x + 1))
                        os.remove(IsItImage(doc))
                req = "DELETE FROM documents WHERE enreg = %s"%(enreg)
                self.c.execute(req)
            dlg = wx.MessageDialog(parent=GLOBVAR.app,
                                   message=mess2,
                                   caption=u"Suppression",
                                   style=wx.OK|wx.ICON_INFORMATION)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.listResultat.DeleteAllItems()
            self.Rechercher()
                
class Classement(wx.SplitterWindow):
    def __init__(self, parent):
        wx.SplitterWindow.__init__(self, parent = parent, id = -1, style=wx.SP_3D)
        l, h = parent.GetClientSizeTuple()
        larg = int(l/3)
        panel1=wx.Panel(self, -1)
        box1 = wx.BoxSizer(wx.VERTICAL)
        self.cla = AClasser(panel1, self)
        box1.Add(self.cla, 1, wx.EXPAND|wx.ALL, border=3)
        panel1.SetSizer(box1)
        box1.Fit(panel1)
        panel1.SetAutoLayout(True)
        panel2=wx.Panel(self, -1)
        box2 = wx.BoxSizer(wx.VERTICAL)
        self.ecran = Affichage(panel2)
        box2.Add(self.ecran, 1, wx.EXPAND|wx.ALL, border=3)
        panel2.SetSizer(box2)
        box2.Fit(panel2)
        panel2.SetAutoLayout(True)
        self.SetMinimumPaneSize(100)
        self.SplitVertically(panel1, panel2, larg)

    def SetImage(self, liste=None):
        if liste:
            self.ecran.SetDocument(liste)
        else:
            self.ecran.SetDocument()
            
class InitScanner(wx.Dialog):
    def __init__(self, titre, liste):
        wx.Dialog.__init__(self, parent=GLOBVAR.app, title = titre)
        self.liste = liste
        choix = []
        box = wx.BoxSizer(wx.VERTICAL)
        mess = wx.StaticText(self, -1, u"Choisissez votre scanner\ndans la liste ci-dessous", style = wx.ALIGN_CENTRE)
        box.Add(mess, 0, wx.EXPAND|wx.ALL, border=5)
        for x in self.liste:
            choix.append("Scanner %s %s"%(x[1], x[2]))
        self.listechoix = wx.Choice(self, -1, choices=choix)
        if SCANVAR.device != "INDEFINI":
            res = self.listechoix.SetStringSelection(SCANVAR.device)
            if not res:
                SCANVAR.SetDevice("INDEFINI")
        box.Add(self.listechoix, 0, wx.EXPAND|wx.ALL, border=5)
        sizerButtons = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        box.Add(sizerButtons, 0, wx.ALIGN_RIGHT|wx.ALL, border=5)
        self.SetSizer(box)
        self.SetAutoLayout(True)
        
        self.Bind(wx.EVT_CHOICE, self.NouveauChoix, self.listechoix)
        
    def NouveauChoix(self, event):
        if self.listechoix.GetStringSelection() != SCANVAR.device:
            SCANVAR.SetDevice(self.listechoix.GetStringSelection())
            SCANVAR.SetMode("INDEFINI")
            SCANVAR.SetResolution("INDEFINI")
            
    def GetValue(self):
        if self.listechoix.GetSelection() == wx.NOT_FOUND:
            return []
        else:
            return self.liste[self.listechoix.GetSelection()]
            
class ParamScanner(wx.Panel):
    def __init__(self, dialog, scanner):
        wx.Panel.__init__(self, parent=dialog, id=-1)
        self.sauve = False
        self.scanner = scanner
        self.par = dialog
        listeModes = self.scanner['mode'].constraint
        listeReso = self.scanner['resolution'].constraint
        self.listeResos = []
        self.choixAutre = False
        for x in listeReso:
            self.listeResos.append(str(x))
        self.listeResos.append("Autre...")
        if SCANVAR.mode != "INDEFINI":
            mode = SCANVAR.mode
            self.scanner.mode = mode
        else:
            self.sauve = True
            mode = self.scanner.mode
            SCANVAR.SetMode(mode)
        if SCANVAR.resolution != "INDEFINI":
            if SCANVAR.resolution not in self.listeResos:
                self.choixAutre = True
            reso = SCANVAR.resolution
            self.scanner.resolution = int(reso)
        else:
            reso = str(self.scanner.resolution)
            SCANVAR.SetResolution(reso)
        self.resoEnCours = reso
        self.im = None
        sizer = wx.BoxSizer(wx.VERTICAL)
        box1 = wx.BoxSizer(wx.HORIZONTAL)
        textMode = wx.StaticText(self, -1, u"Mode de numérisation", style = wx.ALIGN_LEFT)
        self.choixMode = wx.Choice(self, -1, choices=listeModes)
        self.choixMode.SetStringSelection(mode)
        box1.Add(textMode, 2, wx.ALL, border=5)
        box1.Add(self.choixMode, 1, wx.ALL, border=5)
        sizer.Add(box1, 0, wx.EXPAND)
        box2=wx.BoxSizer(wx.HORIZONTAL)
        textReso = wx.StaticText(self, -1, u"Choix de la résolution", style = wx.ALIGN_LEFT)
        self.choixReso = wx.Choice(self, -1, choices=self.listeResos)
        if self.choixAutre:
            self.choixReso.SetStringSelection("Autre...")
        else:
            self.choixReso.SetStringSelection(reso)
        box2.Add(textReso, 2, wx.ALL, border=5)
        box2.Add(self.choixReso, 1, wx.ALL, border=5)
        sizer.Add(box2, 0, wx.EXPAND)
        box3=wx.BoxSizer(wx.HORIZONTAL)
        self.autreReso = wx.StaticText(self, -1, " ", style = wx.ALIGN_LEFT)
        self.valReso = wx.StaticText(self, -1, " ", style = wx.ALIGN_CENTRE)
        self.autreReso.SetForegroundColour(wx.RED)
        self.valReso.SetForegroundColour(wx.RED)
        if self.choixAutre:
            self.autreReso.SetLabel(u"Résolution forcée à :")
            self.valReso.SetLabel(reso)
        #self.entryReso.Enable(False)
        box3.Add(self.autreReso, 2, wx.ALL, border=5)
        box3.Add(self.valReso, 1, wx.ALL, border=5)
        sizer.Add(box3, 0, wx.EXPAND)
        titre2 = wx.StaticText(self, -1, u"Luminosité", style = wx.ALIGN_CENTRE)
        sizer.Add(titre2, 0, wx.EXPAND|wx.ALL, border=5)
        self.chxLum = wx.Slider(self, -1, 10, 0, 20)
        sizer.Add(self.chxLum, 0, wx.EXPAND|wx.ALL, border=5)
        titre3 = wx.StaticText(self, -1, u"Contraste", style = wx.ALIGN_CENTRE)
        sizer.Add(titre3, 0, wx.EXPAND|wx.ALL, border=5)
        self.chxCont = wx.Slider(self, -1, 10, 0, 20)
        sizer.Add(self.chxCont, 0, wx.EXPAND|wx.ALL, border=5)
        titre4 = wx.StaticText(self, -1, u"Précision du détail", style = wx.ALIGN_CENTRE)
        sizer.Add(titre4, 0, wx.EXPAND|wx.ALL, border=5)
        self.chxNet = wx.Slider(self, -1, 10, 0, 20)
        sizer.Add(self.chxNet, 0, wx.EXPAND|wx.ALL, border=5)
        bouton1 = wx.Button(self, -1, u"Aperçu")
        sizer.Add((0,0), 1)
        sizer.Add(bouton1, 0, wx.EXPAND|wx.ALL, border=5)
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        
        self.Bind(wx.EVT_BUTTON, self.Apercu, bouton1)
        self.Bind(wx.EVT_CHOICE, self.ChxMode, self.choixMode)
        self.Bind(wx.EVT_CHOICE, self.ChxReso, self.choixReso)

    def ChxMode(self, event):
        choix = self.choixMode.GetStringSelection()
        self.scanner.mode = choix.encode('utf-8')
        SCANVAR.SetMode(choix)
        self.sauve=True
        self.im = None

    def ChxReso(self, event):
        leChoix = self.choixReso.GetStringSelection()
        if leChoix == "Autre...":
            dlg = wx.TextEntryDialog(GLOBVAR.app,
                                     message = u"Veuillez saisir une valeur de résolution\nvalide pour votre scanner",
                                     caption = u"Forçage de la résolution")
            val = dlg.ShowModal()
            newReso = dlg.GetValue()
            dlg.Destroy()
            if val == wx.ID_OK and newReso:
                try:
                    choix = int(newReso)
                except:
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           caption = u"Opération impossible",
                                           message = u"Vous avez choisi une valeur non numérique",
                                           style = wx.OK|wx.ICON_ERROR)
                    rep = dlg.ShowModal()
                    dlg.Destroy()
                    self.choixReso.SetStringSelection(self.resoEnCours)
                    return
                self.scanner.resolution = choix
                control = str(self.scanner.resolution)
                if control in self.listeResos:
                    self.choixReso.SetStringSelection(control)
                    self.choixAutre = False
                    self.autreReso.SetLabel(" ")
                    self.valReso.SetLabel(" ")
                    self.Layout()
                else:
                    self.choixAutre = True
                    self.autreReso.SetLabel(u"Résolution forcée à :")
                    self.valReso.SetLabel(control)
                    self.Layout()
                SCANVAR.SetResolution(control)
                self.resoEnCours = control
            else:
                self.choixReso.SetStringSelection(self.resoEnCours)
                return
        else:
            if self.choixAutre:
                self.autreReso.SetLabel(" ")
                self.valReso.SetLabel(" ")
                self.choixAutre=False
            choix = int(leChoix)    
            self.resoEnCours = str(choix)
            self.scanner.resolution = choix
            SCANVAR.SetResolution(str(choix))
        self.sauve=True
        self.im = None

    def Apercu(self, event):
        if self.sauve:
            self.sauve = False
            self.file_cfg = os.path.join(GLOBVAR.homedir, ".paprass.cfg")
            config = ConfigParser.ConfigParser()
            config.read([self.file_cfg])
            config.set("scanner", "device", SCANVAR.device)
            config.set("scanner", "mode", SCANVAR.mode)
            config.set("scanner", "resolution", SCANVAR.resolution)
            file_cfg = open(self.file_cfg, 'wb')
            config.write(file_cfg)
            file_cfg.close()
        wx.BeginBusyCursor()
        lum = self.chxLum.GetValue() / 10.
        cont = self.chxCont.GetValue() / 10.
        net = self.chxNet.GetValue() / 10.
        fic = os.path.join(GLOBVAR.tempdir, "apercu.jpg")
        ficTemp = os.path.join(GLOBVAR.tempdir, "test.jpg")
        if not self.im :
            self.scanner.start()
            self.im = self.scanner.snap()
            self.im.save(fic)
        im = Image.open(fic)
        enh = ImageEnhance.Brightness(im)
        im = enh.enhance(lum)
        enh = ImageEnhance.Contrast(im)
        im = enh.enhance(cont)
        enh = ImageEnhance.Sharpness(im)
        im = enh.enhance(net)
        im.save(ficTemp)
        self.par.SetFichier(ficTemp)
        wx.EndBusyCursor()

class NumProcess(wx.Dialog):
    def __init__(self, titre, scanner):
        l, h = wx.ScreenDC().GetSizeTuple()
        larg = (l * 2) / 3
        haut = (h * 2) / 3
        taille = wx.Size(larg, haut)
        wx.Dialog.__init__(self, parent=GLOBVAR.app, title=titre, size=taille)
        self.scanner = scanner
        self.fichier = None
        sizerPrinc = wx.BoxSizer(wx.VERTICAL)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        panneau = ParamScanner(self, self.scanner)
        self.doc = Apercu(self)
        sizer.Add(panneau, 2, wx.EXPAND|wx.ALL, border=5)
        sizer.Add(self.doc, 3, wx.EXPAND|wx.ALL, border=5)
        sizerPrinc.Add(sizer, 1, wx.EXPAND)
        sizerButtons = self.CreateButtonSizer(wx.OK|wx.CANCEL)
        sizerPrinc.Add(sizerButtons, 0, wx.ALIGN_RIGHT|wx.ALL, border=5)
        self.SetSizer(sizerPrinc)
        self.SetAutoLayout(True)

    def SetFichier(self, fichier):
        self.fichier = fichier
        img = wx.Bitmap(fichier, wx.BITMAP_TYPE_JPEG)
        self.doc.SetImage(img)

    def GetFichier(self):
        return self.fichier
        
class Numerisation(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent, id=-1)

        self.serieEnCours = False
        self.liste = []
        self.page = 0
        self.choixPDF = False
        self.txt = False

        box1 = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        self.btCommencer = wx.Button(self, -1, u'Commencer')
        self.Bind(wx.EVT_BUTTON, self.Commencer, self.btCommencer)
        box2.Add(self.btCommencer, 1, wx.ALL, border= 5)

        self.btEnregistrer = wx.Button(self, -1, u'Enregistrer')
        self.Bind(wx.EVT_BUTTON, self.Enreg, self.btEnregistrer)
        box2.Add(self.btEnregistrer, 1, wx.ALL, border= 5)

        self.btAnnuler = wx.Button(self, -1, u'Annuler')
        self.Bind(wx.EVT_BUTTON, self.Annuler, self.btAnnuler)
        box2.Add(self.btAnnuler, 1, wx.ALL, border= 5)

        box1.Add(box2, 0, wx.EXPAND)
        self.panneau = Affichage(self)
        box1.Add(self.panneau, 1, wx.EXPAND)
        self.SetSizer(box1)
        self.SetAutoLayout(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)

    def Annuler(self, event):
        self.panneau.SetDocument()
        self.btCommencer.Enable(True)
        self.btEnregistrer.Enable(False)
        self.btAnnuler.Enable(False)

    def Enreg(self, event):
        self.Enregistrer()

    def Enregistrer(self, pdf = False, path = None, ooo = "non"):
        mode = 3
        dlgTxt = Enregistrer(GLOBVAR.app, mode, path)
        val = dlgTxt.ShowModal()
        resultat1 = dlgTxt.GetTitre()
        resultat2 = dlgTxt.GetDate()
        self.choixPDF, self.A4 = dlgTxt.GetPDF()
        dlgTxt.Destroy()
        if val == wx.ID_OK:
            if resultat1 != "":
                leTitre = "''".join(resultat1.split("'"))
                leTitre=eval('u"%s"'%leTitre)
                laDate = resultat2.FormatISODate()
                lAnnee = laDate.split("-")[0]
                leMois = laDate.split("-")[1]
                if self.choixPDF :
                    nbPages = "1"
                else:
                    nbPages = str(len(self.liste))
                c = GLOBVAR.base.cursor()
                req = "INSERT INTO documents(classeur, dossier, chemise, date, titre, nbpages, annee, mois) "
                req = req + "VALUES(0, 0, 0,'%s', '%s', %s, '%s', '%s')"%(laDate, leTitre, nbPages, lAnnee, leMois)
                c.execute(req)
                res = c.execute("SELECT MAX(enreg) FROM documents")
                numero = str(res.fetchone()[0])
                i = 0
                chem = os.path.join(GLOBVAR.docdir, lAnnee)
                if os.path.isdir(chem)== False:
                    os.mkdir(chem)
                chem = os.path.join(chem, leMois)
                if os.path.isdir(chem) == False :
                    os.mkdir(chem)
                if self.choixPDF :
                    fichier = numero + "-1.pdf"
                    pathComplet = os.path.join(chem, fichier)
                    c = canvas.Canvas(pathComplet)
                    for x in self.liste:
                        im = Image.open(x)
                        largIm, hautIm = im.size
                        if self.A4:
                            c.setPageSize(A4)
                            largFeuille, hautFeuille = A4
                            if largIm > hautIm:
                                im.rotate(90)
                            largIm, hautIm = im.size
                            diff = abs((largFeuille-40) - largIm)
                            if diff != 0:
                                if (largFeuille-40) > largIm :
                                    ajust = (diff / (largFeuille-40)) * 1.
                                    ratio = 1 + ajust
                                else:
                                    ajust = (diff / largIm) * 1.
                                    ratio = 1 - ajust
                            else:
                                ratio = 1
                            larg = int(largIm * ratio)
                            haut = int(hautIm * ratio)
                            taille = (larg, haut)
                            im = im.resize(taille, Image.BICUBIC)
                            c.drawInlineImage(im, 20, (hautFeuille-20) - haut)
                        else:
                            c.setPageSize((largIm, hautIm))
                            c.drawInlineImage(im, 0, 0)
                        c.showPage()
                    c.save()
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document PDF a bien été enregistré",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                else:
                    i = 0
                    for x in self.liste:
                        i += 1
                        fichier = numero + "-" + str(i) + ".jpg"
                        pathComplet = os.path.join(chem, fichier)
                        shutil.copyfile(x, pathComplet)
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Le document et ses images ont été enregistrés",
                                           caption = u"Opération réalisée",
                                           style=wx.OK|wx.ICON_INFORMATION)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                self.panneau.SetDocument()
                self.btCommencer.Enable(True)
                self.btEnregistrer.Enable(False)
                self.btAnnuler.Enable(False)

    def Commencer(self, event=None):
        if not self.serieEnCours:
            sane.init()
            listeScanner = sane.get_devices()
            nbre = len(listeScanner)
            if nbre == 0:
                dlg = wx.MessageDialog(GLOBVAR.app,
                                       message=u"Aucune source Sane reconnue",
                                       caption=u"Impossible de continuer",
                                       style=wx.OK|wx.ICON_ERROR)
                val = dlg.ShowModal()
                dlg.Destroy()
                return
            if nbre == 1 :
                self.titre = u"Scanner %s %s"%(listeScanner[0][1], listeScanner[0][2])
                if SCANVAR.device != "Scanner %s %s"%(listeScanner[0][1], listeScanner[0][2]):
                    SCANVAR.SetDevice("Scanner %s %s"%(listeScanner[0][1], listeScanner[0][2]))
                    SCANVAR.SetMode("INDEFINI")
                    SCANVAR.SetResolution("INDEFINI")
                self.scanner = sane.open(listeScanner[0][0])
                self.page = 0
            else:
                dlg = InitScanner(u"Choix du scanner", listeScanner)
                val = dlg.ShowModal()
                resu=dlg.GetValue()
                dlg.Destroy()
                if val == wx.ID_OK and len(resu) > 0:
                    self.titre = u"Scanner %s %s"%(resu[1], resu[2])
                    self.scanner = sane.open(resu[0])
                    self.page = 0
                else:
                    dlg = wx.MessageDialog(GLOBVAR.app,
                                           message=u"Abandon de l'opération par l'utilisateur",
                                           caption = u"Annulation",
                                           style = wx.OK|wx.ICON_ERROR)
                    val = dlg.ShowModal()
                    dlg.Destroy()
                    return
        else:
            self.page = self.page + 1
        if self.page == 0 :
            self.liste = []
        page = str(self.page)
        if len(page) < 2 :
            page = "0" + page
        self.ficTemp = os.path.join(GLOBVAR.tempdir, "tmp" + page + ".jpg")
        mess = u"Placez votre page %s dans le scanner puis validez"%str(self.page + 1)
        dlg = wx.MessageDialog(GLOBVAR.app,
                               message=mess,
                               caption = u"Numérisation",
                               style=wx.OK|wx.ICON_INFORMATION)
        val = dlg.ShowModal()
        dlg.Destroy()
        dlg = NumProcess(self.titre, self.scanner)
        rep = dlg.ShowModal()
        ficTemp = dlg.GetFichier()
        dlg.Destroy()
        if rep == wx.ID_CANCEL:
            dlg = wx.MessageDialog(GLOBVAR.app,
                                   message=u"Abandon de l'opération par l'utilisateur",
                                   caption = u"Annulation",
                                   style=wx.OK|wx.ICON_ERROR)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.serieEnCours = False
            self.scanner.close()
            return
        elif ficTemp==None:
            dlg = wx.MessageDialog(GLOBVAR.app,
                                   message=u"Vous devez numériser une image avant de valider",
                                   caption = u"Erreur de manipulation",
                                   style=wx.OK|wx.ICON_ERROR)
            val = dlg.ShowModal()
            dlg.Destroy()
            self.page = self.page - 1
            self.Commencer()
            return
        shutil.copyfile(ficTemp, self.ficTemp)
        self.liste.append(self.ficTemp)
        self.panneau.SetDocument(self.liste, page = len(self.liste))
        dlg = wx.MessageDialog(GLOBVAR.app,
                               message = u"Voulez-vous ajouter une page à votre document ?",
                               caption = u"Numérisation",
                               style=wx.YES_NO|wx.ICON_QUESTION)
        val = dlg.ShowModal()
        dlg.Destroy()
        if val == wx.ID_YES:
            self.serieEnCours = True
            self.Commencer()
        else:
            self.serieEnCours = False
            self.scanner.close()
        if len(self.liste) != 0:
            self.btCommencer.Enable(False)
            self.btEnregistrer.Enable(True)
            self.btAnnuler.Enable(True)

class Principale(wx.Frame):
    def __init__(self, titre):
        largE, hautE = wx.ScreenDC().GetSizeTuple()
        larg = (largE * 5) / 6
        haut = (hautE * 5) / 6
        taille = wx.Size(larg, haut)
        wx.Frame.__init__(self, parent = None, id = -1, title = titre, size=taille)
        chemImage = os.path.join(GLOBVAR.themedir, "bipede.png")
        self.SetIcon(wx.Icon(chemImage, wx.BITMAP_TYPE_PNG))
        self.menuDoc = wx.Menu()
        self.menuClass = wx.Menu()
        self.menuAide = wx.Menu()
        menuBar = wx.MenuBar()

        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Numériser", help = u"Numériser un nouveau document")
        self.ID_NUM = item.GetId()
        self.menuDoc.AppendItem(item)
        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Ajouter", help = u"Ajouter un document existant")
        self.ID_AJOUT = item.GetId()
        self.menuDoc.AppendItem(item)
        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Ajouter une note", help = u"Ajouter une note manuscrite")
        self.ID_NOTE = item.GetId()
        self.menuDoc.AppendItem(item)
        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Fermer", help = u"Fermer le module en cours")
        self.ID_FERMER = item.GetId()
        self.menuDoc.AppendItem(item)
        self.menuDoc.AppendSeparator()
        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Rechercher", help = u"Rechercher un document")
        self.ID_RECH = item.GetId()
        self.menuDoc.AppendItem(item)
        self.menuDoc.AppendSeparator()
        item = wx.MenuItem(self.menuDoc, id = -1, text = u"Quitter", help = u"Quitter Pap'rass")
        self.ID_QUIT = item.GetId()
        self.menuDoc.AppendItem(item)
        item = wx.MenuItem(self.menuClass, id = -1, text = u"Configurer", help = u"Configurer Pap'rass")
        self.ID_DEF = item.GetId()
        self.menuClass.AppendItem(item)
        item = wx.MenuItem(self.menuClass, id = -1, text = u"Classer", help = u"Classer les documents dans leur chemise")
        self.ID_INDEX = item.GetId()
        self.menuClass.AppendItem(item)
        item = wx.MenuItem(self.menuAide, id = -1, text = u"A propos", help = u"A propos de Pap'rass")
        self.ID_AIDE = item.GetId()
        self.menuAide.AppendItem(item)

        menuBar.Append(self.menuDoc, u"Documents")
        menuBar.Append(self.menuClass, u"Classement")
        menuBar.Append(self.menuAide,u"?")

        self.SetMenuBar(menuBar)

        self.barre = wx.ToolBar(self, -1, style=wx.TB_TEXT)

        self.barre.SetToolBitmapSize((48, 48))

        self.btNumeriser = self.barre.AddLabelTool(self.ID_NUM, u"Numériser",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "numeriser.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Numériser un nouveau document")
        self.btAjoutDoc = self.barre.AddLabelTool(self.ID_AJOUT, u"Ajouter",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "ajouter.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Ajouter un document existant")
        self.btAjoutNote = self.barre.AddLabelTool(self.ID_NOTE, u"Note",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "note.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Ajouter une note manuscrite")
        self.barre.AddSeparator()
        self.btIndexer = self.barre.AddLabelTool(self.ID_INDEX, u"Classer",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "classer.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Classer les documents dans leur chemise")
        self.btRechercher = self.barre.AddLabelTool(self.ID_RECH, u"Rechercher",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "rechercher.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Rechercher un document")
        self.barre.AddSeparator()
        self.btClassement = self.barre.AddLabelTool(self.ID_DEF, u"Configurer",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "config.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Configurer Pap'rass")
        self.btFermer = self.barre.AddLabelTool(self.ID_FERMER, u"Fermer",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "fermer.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Fermer le module en cours")
        self.barre.AddSeparator()
        self.btQuitter = self.barre.AddLabelTool(self.ID_QUIT, u"Quitter",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "quitter.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"Quitter l'application")
        self.btAPropos = self.barre.AddLabelTool(self.ID_AIDE, "A propos",
                        wx.Bitmap(os.path.join(GLOBVAR.themedir, "aide.png"), wx.BITMAP_TYPE_PNG),
                        shortHelp = u"A propos de Pap'rass...")

        self.barre.Realize()
        self.SetToolBar(self.barre)
        statBar = wx.StatusBar(self)
        self.SetStatusBar(statBar)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.conteneur = wx.Panel(self, -1)
        sizer.Add(self.conteneur, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.conteneurSizer = wx.BoxSizer(wx.VERTICAL)
        self.panneau = PanneauLogo(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.conteneur.SetSizer(self.conteneurSizer)
        self.conteneur.SetAutoLayout(True)
        self.SetAutoLayout(True)
        self.Centre()

        self.barre.EnableTool(self.ID_FERMER, False)
        self.menuDoc.Enable(self.ID_FERMER, False)
        if WIN:
            self.barre.EnableTool(self.ID_NUM, False)
            self.menuDoc.Enable(self.ID_NUM, False)

        wx.EVT_MENU(self, self.ID_NUM, self.Numeriser)
        wx.EVT_MENU(self, self.ID_AJOUT, self.AjoutDoc)
        wx.EVT_MENU(self, self.ID_NOTE, self.AjoutNote)
        wx.EVT_MENU(self, self.ID_INDEX, self.Classer)
        wx.EVT_MENU(self, self.ID_RECH, self.Rechercher)
        wx.EVT_MENU(self, self.ID_DEF, self.Configurer)
        wx.EVT_MENU(self, self.ID_FERMER, self.Fermer)
        wx.EVT_MENU(self, self.ID_QUIT, self.OnClose)
        wx.EVT_MENU(self, self.ID_AIDE, self.APropos)

        wx.EVT_CLOSE(self, self.OnClose)

        wx.EVT_SIZE(self, self.Rafraichir)

    def Rafraichir(self, event):
        self.conteneur.SetSize(self.GetClientSize())
        self.conteneur.SetPosition((0, 0))
        self.panneau.SetSize(self.conteneur.GetClientSize())
        self.panneau.Layout()

    def Neutraliser(self):
        self.barre.EnableTool(self.ID_NUM, False)
        self.menuDoc.Enable(self.ID_NUM, False)
        self.barre.EnableTool(self.ID_AJOUT, False)
        self.menuDoc.Enable(self.ID_AJOUT, False)
        self.barre.EnableTool(self.ID_NOTE, False)
        self.menuDoc.Enable(self.ID_NOTE, False)
        self.barre.EnableTool(self.ID_INDEX, False)
        self.menuClass.Enable(self.ID_INDEX, False)
        self.barre.EnableTool(self.ID_RECH, False)
        self.menuDoc.Enable(self.ID_RECH, False)
        self.barre.EnableTool(self.ID_DEF, False)
        self.menuClass.Enable(self.ID_DEF, False)
        self.barre.EnableTool(self.ID_FERMER, True)
        self.menuDoc.Enable(self.ID_FERMER, True)
        self.barre.EnableTool(self.ID_QUIT, False)
        self.menuDoc.Enable(self.ID_QUIT, False)

    def Numeriser(self, evt):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Numériser un document")
        self.panneau = Numerisation(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def AjoutDoc(self, event):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Ajouter un document")
        self.panneau = AjoutFichier(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def AjoutNote(self, event):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Ajouter une note manuscrite")
        self.panneau = AjoutNote(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def Classer(self, event):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Classer")
        self.panneau = Classement(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def Rechercher(self, event):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Rechercher")
        self.panneau = Recherche(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def Configurer(self, event):
        self.Neutraliser()
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass - Configurer Pap'rass")
        self.panneau = Configuration(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        self.SendSizeEvent()

    def Fermer(self, event):
        self.panneau.Destroy()
        self.SetTitle(u"Pap'rass")
        self.panneau = PanneauLogo(self.conteneur)
        self.conteneurSizer.Add(self.panneau, 1, wx.EXPAND)
        if not WIN:
            self.barre.EnableTool(self.ID_NUM, True)
            self.menuDoc.Enable(self.ID_NUM, True)
        self.barre.EnableTool(self.ID_AJOUT, True)
        self.menuDoc.Enable(self.ID_AJOUT, True)
        self.barre.EnableTool(self.ID_NOTE, True)
        self.menuDoc.Enable(self.ID_NOTE, True)
        self.barre.EnableTool(self.ID_INDEX, True)
        self.menuClass.Enable(self.ID_INDEX, True)
        self.barre.EnableTool(self.ID_RECH, True)
        self.menuDoc.Enable(self.ID_RECH, True)
        self.barre.EnableTool(self.ID_DEF, True)
        self.menuClass.Enable(self.ID_DEF, True)
        self.barre.EnableTool(self.ID_FERMER, False)
        self.menuDoc.Enable(self.ID_FERMER, False)
        self.barre.EnableTool(self.ID_QUIT, True)
        self.menuDoc.Enable(self.ID_QUIT, True)
        self.SendSizeEvent()

    def APropos(self, event):
        dlg = APropos()
        rep = dlg.ShowModal()
        dlg.Destroy()

    def OnClose(self, event):
        GLOBVAR.base.close()
        self.Destroy()

class Paprass(wx.App):
    def OnInit(self):
        wx.InitAllImageHandlers()
        f = Principale("Pap'rass")
        GLOBVAR.app = f
        f.Show(True)
        self.SetTopWindow(f)
        return True

GLOBVAR = GlobalVar()
CONFIG = Config()
SCANVAR = ScanVar()
GLOBVAR.themedir = os.path.join("/usr/share/paprass/themes", CONFIG.GetTheme())
GLOBVAR.visupdf = CONFIG.GetViewer()
GLOBVAR.base = CONFIG.GetBase()
SCANVAR.SetDevice(CONFIG.GetDevice())
SCANVAR.SetMode(CONFIG.GetMode())
SCANVAR.SetResolution(CONFIG.GetResolution())

app = Paprass()
app.MainLoop()
