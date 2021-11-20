import datetime
import pandas as pd
from PyQt6 import uic
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog

import sys

if not sys.warnoptions:
    import warnings
    warnings.simplefilter("ignore")

qtCreatorFile = "app.ui"  # Aquí va el nombre de tu archivo

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)
print(uic.loadUiType(qtCreatorFile))


class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.path = None
        self.d_CnMaestro = None
        self.d_SM = None
        self.d_Conectantes = None
        self.d_FaseDDA = None

        #--------------------- Botones -----------------------------------------------
        self.CnButton.clicked.connect(self.getExcel)
        self.SMButton.clicked.connect(self.getExcel2)
        self.crearButton.clicked.connect(self.crearArchivo)
        # -----------------------------------------------------------------------------


    def CreateDataframe(self,Est_Aps, Df_SM, Df_Conectante, Df_FasesDDA,num,devices):
        print("Entro función create\n")


        print(Df_FasesDDA.columns)
        print(Df_Conectante.columns)
        Est_Aps = Est_Aps.reset_index()

        filtro = Df_SM[Df_SM.ID_BENEFICIARIO.isin(Est_Aps.ID_BENEFICIARIO)]

        Est_Aps = filtro[['ID', 'ID_BENEFICIARIO', "IM", "ESTADO", "PRIORIDAD", "TÍTULO", "FECHA_HORA_DE_APERTURA"]]
        #print(Est_Aps.head(10))

        filtro2 = Df_Conectante[Df_Conectante.ID.isin(Est_Aps.ID)]
        #print(filtro2.columns)
        filtro3 = Df_FasesDDA[Df_FasesDDA.ID.isin(Est_Aps.ID)]
        #print(filtro3.columns)
        fases_DDA = filtro3[["ID", "MUNICIPIO", "FASE_OFICIAL", "DDA"]]
        bts = filtro2[['ID', "BTS"]]
        #print(bts)
        id_bts = pd.merge(bts, Est_Aps, on='ID')

        Pag = pd.merge(id_bts, fases_DDA, on='ID', how='outer')

        if num == 1:

            Pag = pd.merge(devices, Pag, on='ID_BENEFICIARIO')
            Pag = Pag.drop(columns=["Online", "Offline"], axis=1)


            Pag["DIAGNOSTICO"] = "Se realiza la validacion del CD para el ID BENEFICIARIO: " + Pag[
                "ID_BENEFICIARIO"].astype(str) + ", ubicado en el MUNICIPIO: " \
                                 + Pag["MUNICIPIO"].astype(str) + ", con DDA: " + Pag["DDA"].astype(
                str) + " encontrando falla en el AP: " + Pag["DEVICE_NAME"].astype(
                str) + " por lo cual es necesario generar Tarea para desplazar personal a sitio y realizar las respectivas validaciones." \
                                 + "Nota: Se solicita al personal que se va a desplazar a sitio llevar repuestos para las APS, Tarjeta de Red para la UPS,Cable UTP Categoria 6, Garantizar que el PC soporte troncalizacion de Vlan y funcione correctamente la tarjeta de RED."

            Pag = Pag.groupby(['ID', 'ID_BENEFICIARIO', 'BTS', 'IM', 'ESTADO',
                               'PRIORIDAD', "TÍTULO", "FECHA_HORA_DE_APERTURA", 'MUNICIPIO', 'FASE_OFICIAL', 'DDA',
                               "DIAGNOSTICO"])['DEVICE_NAME'].apply(list)


        elif num == 2:

            Pag["DIAGNOSTICO"] = "Se realiza la validacion de la Ruta de Tx,para el ID BENEFICIARIO: " + Pag[
                "ID_BENEFICIARIO"] + ", ubicado en el MUNICIPIO: " + Pag["MUNICIPIO"] + ", con DDA: " + Pag[
                                     "DDA"] + ", donde se descartan fallas  en el los distintos tramos de la Ruta de Tx" \
                                 + "por lo cual es necesario generar Tarea para desplazar personal a sitio y realizar las respectivas validaciones."

            Pag = Pag.groupby(
                ['ID', 'BTS', 'ID_BENEFICIARIO', 'IM', "ESTADO", "PRIORIDAD", "TÍTULO", "FECHA_HORA_DE_APERTURA",
                 'MUNICIPIO', 'FASE_OFICIAL', 'DIAGNOSTICO'])['DDA'].apply(list)
        else:

            Pag["DIAGNOSTICO"] = "Se realiza el monitorieo para el ID BENEFICIARIO: " + Pag[
                "ID_BENEFICIARIO"] + ", ubicado en el MUNICIPIO: " + Pag["MUNICIPIO"] + ", con DDA: " + Pag[
                                     "DDA"] + ", donde se evidencia que el CD se encuentra operativo, y todos los APs se encuentran Online."

            Pag = Pag.groupby(
                ['ID', 'BTS', 'ID_BENEFICIARIO', 'IM', "ESTADO", "PRIORIDAD", "TÍTULO", "FECHA_HORA_DE_APERTURA",
                 'MUNICIPIO', 'FASE_OFICIAL', 'DIAGNOSTICO'])['DDA'].apply(list)
        print(Pag)
        print("/////////__________________________________Finish function Create__________________________________/////////\n")
        return Pag

    def analisys_data(self,path):
        self.d_Conectantes = pd.read_excel("Libro1.xlsx")
        self.d_FaseDDA = pd.read_excel('Fase1A_1B_conDDA.xlsx')
        d_CnMaestro = self.d_CnMaestro
        d_SM = self.d_SM
        d_FaseDDA = self.d_FaseDDA
        d_Conectantes = self.d_Conectantes

        print("Setup complete.")
        #______________________________________________________________________________________________________________________
        d_CnMaestro.columns = map(str.upper, d_CnMaestro.columns)
        d_SM.columns = map(str.upper, d_SM.columns)
        d_FaseDDA.columns = map(str.upper, d_FaseDDA.columns)
        d_Conectantes.columns = map(str.upper, d_Conectantes.columns)

        d_SM.columns = d_SM.columns.astype(str).str.replace("[()]", "")
        d_SM.columns = d_SM.columns.astype(str).str.replace(" ", "_")
        d_SM.columns = d_SM.columns.astype(str).str.replace("/", "_")

        d_CnMaestro.columns = d_CnMaestro.columns.astype(str).str.replace("[()]", "")
        d_CnMaestro.columns = d_CnMaestro.columns.astype(str).str.replace(" ", "_")
        d_CnMaestro.columns = d_CnMaestro.columns.astype(str).str.replace("/", "_")

        d_FaseDDA.columns = d_FaseDDA.columns.astype(str).str.replace("[()]", "")
        d_FaseDDA.columns = d_FaseDDA.columns.astype(str).str.replace(" ", "_")
        d_FaseDDA.columns = d_FaseDDA.columns.astype(str).str.replace("/", "_")

        d_Conectantes.columns = d_Conectantes.columns.astype(str).str.replace("[()]", "")
        d_Conectantes.columns = d_Conectantes.columns.astype(str).str.replace(" ", "_")
        d_Conectantes.columns = d_Conectantes.columns.astype(str).str.replace("/", "_")

        array_delete = d_CnMaestro.loc[d_CnMaestro['SITE'] == '777-PILOTO'].index.to_numpy()
        d_CnMaestro = d_CnMaestro.drop(array_delete)

        d_CnMaestro["ID_BENEFICIARIO"] = d_CnMaestro["SITE"].str.extract(r"(\d{5})")

        d_SM = d_SM.rename(columns={"ID_DE_INCIDENTE": "IM"})
        d_SM = d_SM.rename(columns={"ID_MINTIC": "ID"})
        d_SM = d_SM[d_SM.ID_BENEFICIARIO.notna()]
        #
        d_SM['FECHA_HORA_DE_APERTURA'] = pd.to_datetime(d_SM['FECHA_HORA_DE_APERTURA'], infer_datetime_format=True)
        d_SM = d_SM.drop_duplicates()

        d_SM['IM'] = d_SM['IM'].astype(str)
        d_SM['ID'] = d_SM['ID'].astype(str)
        #
        d_SM = d_SM[d_SM['ASIGNADO_A'] == 'Carlos Albeiro. Diaz Tangarife']
        d_SM['ID'] = d_SM['ID'].str.replace("\t", "")
        d_SM = d_SM[d_SM['ID'].str.len() <= 8]
        d_SM = d_SM[d_SM['FECHA_HORA_DE_APERTURA'] >= '2021-09-25']


        d_FaseDDA = d_FaseDDA.rename(columns={"ID_MINTIC": "ID"})
        d_FaseDDA['ID'] = d_FaseDDA['ID'].astype(str)
        d_FaseDDA['ID_BENEFICIARIO'] = d_FaseDDA['ID_BENEFICIARIO'].astype(str)



        d_SM['ID_BENEFICIARIO'] = d_SM['ID_BENEFICIARIO'].astype(str)
        d_Conectantes['ID'] = d_Conectantes['ID'].astype(str)

        filtro1 = d_CnMaestro[d_CnMaestro.ID_BENEFICIARIO.isin(d_SM.ID_BENEFICIARIO)]
        d_Aps = filtro1.groupby(['ID_BENEFICIARIO', "DEVICE_NAME"])['STATUS'].value_counts().unstack().fillna(0)

        d_Aps_1 = d_Aps

        d_Aps = filtro1.groupby(['ID_BENEFICIARIO'])['STATUS'].value_counts().unstack().fillna(0)



        Aps_OnOff = d_Aps.loc[(d_Aps['Online'] > 0.0)]
        Aps_OnOff = Aps_OnOff.loc[(Aps_OnOff['Online'] < 3.0)]
        aps_revisar = Aps_OnOff.reset_index()
        aps_revisar = d_CnMaestro[d_CnMaestro.ID_BENEFICIARIO.isin(aps_revisar.ID_BENEFICIARIO)]
        aps_revisar = aps_revisar.groupby(['ID_BENEFICIARIO', "DEVICE_NAME"])['STATUS'].value_counts().unstack().fillna(0)
        aps_revisar = aps_revisar.reset_index()
        aps_revisar = aps_revisar.loc[aps_revisar['Offline'] > 0.0]


        print("\n ____________________________#$#$#$$#$#$#$#$#$#$#$#$_______________________________\n")

        Aps_Offline = d_Aps.loc[d_Aps['Online'] == 0.0]
        Aps_Online = d_Aps.loc[d_Aps['Online'] == 3.0]
        #
        #
        #
        df_centrodigital = self.CreateDataframe(Aps_OnOff, d_SM, d_Conectantes, d_FaseDDA, 1, aps_revisar)
        df_rutaTx = self.CreateDataframe(Aps_Offline, d_SM, d_Conectantes, d_FaseDDA, 2, aps_revisar)
        df_allonline = self.CreateDataframe(Aps_Online, d_SM, d_Conectantes, d_FaseDDA, 3, aps_revisar)
        print("\n \t>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Finish <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< \n")
        self.to_excel_sheet(path,d_Aps_1, df_rutaTx, df_centrodigital, df_allonline)

    def crearArchivo(self):
        try:
            #path = easygui.diropenbox()
            path = QFileDialog.getExistingDirectory(self, 'Select Folder')
            #path = path.replace(chr(92), '/')
            print(path)
            #QMessageBox.about(self, "Directorio", path)
            #if path != "":

            self.analisys_data(path)

        except Exception as e:
            QMessageBox.about(self, "Mensaje", "Seleccione una carpeta válida")


    def to_excel_sheet(self, path, df_Aps, df_Tx, df_CDs, df_allon):

        print("Entro funcion excel")
        f = datetime.datetime.today().__str__()
        f = f.replace(":", "-")
        f = f.replace(".", "-")
        print(path + "/APsStatus-" + f + ".xlsx")
        #print(df_Tx.columns)

        with pd.ExcelWriter(path + "/APsStatus-" + f + ".xlsx") as writer:
            df_Aps.to_excel(writer, sheet_name="Aps_Status")
            df_Tx.to_excel(writer, sheet_name="Aps_Revisar_TX")
            df_CDs.to_excel(writer, sheet_name="Aps_Revisar_CDs")
            df_allon.to_excel(writer, sheet_name="Aps_AllOnline")
            QMessageBox.about(self, "Mensaje", "Excel creado con éxito")


    def getExcel(self):
        print("entro")
        #print(self.d_Conectantes)
        filePath = QFileDialog.getOpenFileName()
        #print(filePath[0])
        if filePath[0] != "":
            path = str(filePath[0])
            path = path.replace(chr(92), '/')
            QMessageBox.about(self, "Mensaje", "Espere por favor")
            print(path)
            # exc_file = pd.ExcelFile("C:/Users/Admin/OneDrive/ProyectoMinTik7K/03-nov/CN_Maestro.xlsx")
            try:
                self.d_CnMaestro = pd.read_excel(path)
                print(self.d_CnMaestro.head(5))
                QMessageBox.about(self, "Mensaje", "Excel importado con éxito")
            except Exception as e:
                QMessageBox.about(self, "Mensaje", "Seleccione archivo válido, CN Maestro")

    def getExcel2(self):
        print("entro")
        filePath = QFileDialog.getOpenFileName()
        if filePath[0] != "":
            print("Dirección", filePath)  # Opcional imprimir la dirección del archivo
            path = str(filePath[0])
            path = path.replace(chr(92), '/')
            print(path)
            try:
                self.d_SM = pd.read_excel(path,0)
                print(self.d_SM.head(5))
                QMessageBox.about(self, "Mensaje", "Excel importado con éxito")
            except Exception as e:
                QMessageBox.about(self, "Mensaje", "Seleccione archivo válido, Service Manager")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.setWindowTitle("Status Aps - Incidentes")
    window.setWindowIcon(QIcon('icon.png'))

    window.show()
    try:
        sys.exit(app.exec())
    except SystemError:
        print("Closing")