import cx_Oracle
import datetime
from error import Error
import re
import time



class CertProf(object):
    """ Objeto de carga para las filas de los excel de certificados digitales """

    sol_cod = ''
    d_excel = ''          #varchar2(100) not 'None',
    d_pestana = ''       #varchar2(20) not 'None',
    n_fila = ''          #number not 'None',
    expediente = ''      # Desestructuramos el exp
    expFamilia = ''
    expProvincia = ''
    expAno = ''
    expNum = ''
    exp_cod = ''         # Sale del expediente
    exp_prv = ''         # Sale del expediente
    exp_ano = ''         # Sale del expediente
    exp_estatal = ''     # Sale del expediente
    c_estado = ''        # Lo dará la funcion get_estado() en función de indicaciones de Ana Celia
    f_estado = time.strftime("%d/%m/%y")        # Fecha del día de la migración
    c_tipo_acreditacion = '' #Tipo de acreditación en caso positivo. TOTAL, PARCIAL o ANOTACIÓN
    f_present = ''       #date,
    d_dni = ''           #varchar2(9),
    d_apel1 = ''         #varchar2(40),
    d_apel2 = ''         #varchar2(40),
    d_nombre = ''        #varchar2(40),
    c_sexo = ''          #char(1),
    f_nacim = ''         #date,
    d_direccion = ''     #varchar2(100),
    c_postal = ''        #varchar2(5),
    c_municipio = ''
    d_localidad = ''     #varchar2(50),
    d_provincia = ''     #varchar2(40),
    c_via_exped = ''     #number,
    c_certificado = ''   #varchar2(10),
    d_real_decreto = ''  #varchar2(100),
    d_certificado = ''   #varchar2(100),
    f_expedicion = ''    #date,
    d_fcs = ''           #varchar2(100),
    d_mods = ''          #varchar2(40),
    d_doc_present = ''   #varchar2(40),
    d_observaciones = '' #varchar2(512),
    # Aqui en medio aparece una columna con el nombre COD CERTIF. Y/O UNIDADES DE COMPETENCIA ????
    d_cod_certificado = '',
    d_tecn_prop_reg = '' #varchar2(10),
    d_doc_req = ''       #varchar2(100),
    f_doc_req = ''       #date,
    f_doc_pres = ''      #date,
    y_req_bocyl = ''     # S/N
    f_propuesta = ''     #date,
    c_prop_aprob = ''    #char(1),
    c_prop_deneg = ''    #char(1),
    f_resol_deneg = ''   #date,
    f_notif_deneg = ''   #date,
    f_resolucion = ''    # Es f_resol_deneg o f_resol_prov la que tenga valor
    c_registro = ''      #varchar2(11),
    f_resol_prov = ''    #date,
    f_recep_titulo = ''  #date,
    f_envio_titulo = ''  #date,
    f_entrega_titulo = '' #date,
    f_recibi_titulo = ''  #date
    d_signatura_arch = ''
    f_res_informe = ''
    estado = ''
    itinerario = ''
    per_fo_cod = ''
    per_num_doc = ''
    per_cod = ''
    per_tip_doc = ''
    per_let_doc = ''
    exp_fam = ''
    fecha_notificacion_boe = ''
    f_desestimiento = ''
    f_envio_otra_ca = ''
    f_certificacion = ''
    c_res_informe = ''
    c_tipo_solicitud = ''
    uc_insert = ''
    mod_insert = ''
    check_archivado = 'N'

    flagUnidadesFormativas = False
    flagOtrosPresentada = False
    flagOtrosRequerida = False

    arrayOtros = []
    arrayDenegada = []       #TODO: Limipar
    arrayAprobada = []       #TODO: Limipar
    arrayUC = []             #TODO: Limipar
    arrayMod = []            #TODO: Limipar
    arrayDocPresentada = []  #TODO: Limpiar
    arrayEspecialidad = []
    arrayEspecialidadPracticas = []
    arrayLocalidades = []
    arrayLocalidadesUnicas = []

    def set_dni(self, d_dni):
        if not d_dni:
            self.d_dni = ""
            error = Error()
            error.description = str(self.d_excel) + " Falta el DNI"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_dni"
            error.log_error()
        else:
            self.d_dni = d_dni.strip(" ")
            if len(self.d_dni) <= 9:
                if self.validarDNI(d_dni):
                    self.d_dni = d_dni
                else:
                    error = Error()
                    error.description = str(self.d_excel) + " El DNI no es correcto "
                    error.excel_name = self.d_excel
                    error.row_number = self.n_fila
                    error.function = "CertProf.set_dni"
                    error.log_error()
            else:
                self.d_dni = d_dni[0:9]
                error = Error()
                error.description = str(self.d_excel) + " El DNI no contiene 9 digitos, se guardan los 9 primeros"
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.set_dni"
                error.log_error()

    def set_sexo(self, c_sexo):
        if c_sexo:
            c_sexo = c_sexo.strip()
        options = ('H', 'M')
        if c_sexo in options:
            if c_sexo == 'H':
                self.c_sexo = '1'
            else:
                self.c_sexo = '2'
        else:
            error = Error()
            if not c_sexo:
                error.description = str(self.d_excel) + " El campo sexo esta vacío, se coloca 1- Hombre como valor"
            else:
                error.description = str(self.d_excel) + " Sexo no esta en H o M (" + c_sexo + "), se coloca 1 - Hombre como valor"

            self.c_sexo = '1'

            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_sexo"
            error.log_error()

    def set_cp(self, c_postal):
        if not c_postal:
            self.c_postal = ""
            error = Error()
            error.description = str(self.d_excel) + " Falta el código postal"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_cp"
            error.log_error()
        else:
            pattern = re.compile(r'^\d{5}$')
            if pattern.search(c_postal):
                self.c_postal = c_postal
            else:
                self.c_postal = c_postal[0:4]
                error = Error()
                error.description = str(self.d_excel) + " El codigo postal no contiene 5 digitos, se guardan los 5 primeros"
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.set_cp"
                error.log_error()

    def set_localidad(self, d_localidad, cursor):
        if not d_localidad:
            self.d_localidad = ""
            error = Error()
            error.description = str(self.d_excel) + " La localidad esta vacia"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_localidad"
            error.log_error()
        else:
            self.d_localidad = str(d_localidad).strip()
            sql = "sql busca localidad" + self.d_localidad + "'"
            cursor.execute(sql)
            row = cursor.fetchone()
            if row is None:
                # error = Error()
                # error.description = str(
                #     self.d_excel) + " No se encuentra la localidad ( " + str(
                #     self.d_localidad) + " ) Sentencia :" + sql
                # error.excel_name = self.d_excel
                # error.row_number = self.n_fila
                # error.function = "CertProf.set_localidad"
                # error.log_error()
                self.arrayLocalidades.append(str(
                    self.d_localidad) + " " + str(self.exp_fam) + "/" + str(self.exp_ano) + "/" + str(self.exp_prv) + "/" + str(self.exp_cod))
                if self.d_localidad not in self.arrayLocalidadesUnicas:
                    self.arrayLocalidadesUnicas.append(self.d_localidad)
                self.d_localidad = ''
            else:
                # error = Error()
                # error.description = str(
                #     self.d_excel) + " Set localidad funciona con localidad ( " + str(
                #     self.d_localidad) + " ) Sentencia :" + sql
                # error.excel_name = self.d_excel
                # error.row_number = self.n_fila
                # error.function = "CertProf.set_localidad"
                # error.log_error()
                self.d_localidad = str(row[0]).strip()

    def set_provincia(self, d_provincia):
        if not d_provincia:
            self.d_provincia = ""
            error = Error()
            error.description = str(self.d_excel) + " La provincia esta vacia"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_localidad"
            error.log_error()
        else:
            self.d_provincia = d_provincia.strip()

    def set_c_certificado(self, c_certificado):
        c_certificado = str(c_certificado).strip()
        if (not c_certificado) or c_certificado == 'None' or c_certificado is None:
            error = Error()
            error.description = str(self.d_excel) + " El codigo de certificado esta vacio " + c_certificado
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_c_certificado"
            error.log_error()
            self.c_certificado = ""
        else:
            self.c_certificado = c_certificado

    def normalize_dates(self):
        if isinstance(self.f_res_informe, datetime.datetime):
            self.f_res_informe = self.f_res_informe.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_res_informe:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de resultado de informe no tiene formato de fecha REVISAR " + str(self.f_res_informe)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        if isinstance(self.f_present, datetime.datetime):
            self.f_present = self.f_present.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_present:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de presntacion no tiene formato de fecha REVISAR " + str(self.f_present)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de presntacion ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                # error.log_error() TODO: Dejar como error o no ¿?
                # TODO: Poner que vaya en blanco no None

        if isinstance(self.f_resolucion, datetime.datetime):
            self.f_resolucion = self.f_resolucion.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_resolucion:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de resolución no tiene formato de fecha REVISAR " + str(self.f_resolucion)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de resolucion ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"

        self.f_recep_titulo = str(self.f_recep_titulo).replace('BOE', '')
        self.f_recep_titulo = str(self.f_recep_titulo).replace('None', '')
        self.f_recep_titulo = str(self.f_recep_titulo).replace('/', '-')
        self.f_recep_titulo = str(self.f_recep_titulo).replace('00:00:00', '').strip()
        self.f_recep_titulo = str(self.f_recep_titulo).replace('00-00-0000', '').strip()
        if str(self.f_recep_titulo).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"
        if self.f_recep_titulo != '':
            self.f_recep_titulo = datetime.datetime.strptime(str(self.f_recep_titulo), format_fecha)

        if isinstance(self.f_recep_titulo, datetime.datetime):
            self.f_recep_titulo = self.f_recep_titulo.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_recep_titulo:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de recepcion de titulo no tiene formato de fecha REVISAR " + str(self.f_recep_titulo)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de recepcion ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                # error.log_error() TODO: Dejar como error o no ¿?

        if isinstance(self.f_nacim, datetime.datetime):
            self.f_nacim = self.f_nacim.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_nacim:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_nacim no tiene formato de fecha REVISAR " + str(self.f_nacim)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        if isinstance(self.f_propuesta, datetime.datetime):
            self.f_propuesta = self.f_propuesta.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_propuesta:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_propuesta no tiene formato de fecha REVISAR " + str(self.f_propuesta)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        self.f_doc_pres = str(self.f_doc_pres).replace('BOE', '')
        self.f_doc_pres = str(self.f_doc_pres).replace('None', '')
        self.f_doc_pres = str(self.f_doc_pres).replace('/', '-')
        self.f_doc_pres = str(self.f_doc_pres).replace('00:00:00', '').strip()
        if str(self.f_doc_pres).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"
        if self.f_doc_pres != '':
            self.f_doc_pres = datetime.datetime.strptime(str(self.f_doc_pres), format_fecha)

        if isinstance(self.f_doc_pres, datetime.datetime):
            self.f_doc_pres = self.f_doc_pres.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_doc_pres:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_doc_pres no tiene formato de fecha REVISAR " + str(self.f_doc_pres)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        # Tratamos el tema BOE
        if self.f_doc_req and type(self.f_doc_req).__name__ != 'datetime':
            if 'BOE' in str(self.f_doc_req) or 'boe' in str(self.f_doc_req):
                self.f_doc_req = str(self.f_doc_req).replace('BOE', '')
                self.f_doc_req = str(self.f_doc_req).replace('boe', '')
                self.y_doc_req = 'S'
            else:
                self.y_doc_req = 'N'
        else:
            self.y_doc_req = 'N'

        self.f_doc_req = str(self.f_doc_req).replace('None', '')
        self.f_doc_req = str(self.f_doc_req).replace('/', '-')
        self.f_doc_req = str(self.f_doc_req).replace('00:00:00', '').strip()
        self.f_doc_req = str(self.f_doc_req).replace('00-00-0000', '').strip()
        if str(self.f_doc_req).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"

        try:
            if self.f_doc_req != '':
                self.f_doc_req = datetime.datetime.strptime(str(self.f_doc_req), format_fecha)
        except Exception as e:
            error = Error()
            error.description = str(
                self.d_excel) + " Se ha producido un error en la fecha " + str(
                self.f_entrega_titulo) + " -- " + e.__str__()
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.normalize_dates"
            # error.log_error()

        if isinstance(self.f_doc_req, datetime.datetime):
            self.f_doc_req = self.f_doc_req.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_doc_req:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_doc_req no tiene formato de fecha REVISAR " + str(self.f_doc_req)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        self.f_expedicion = str(self.f_expedicion).replace('None', '')
        self.f_expedicion = str(self.f_expedicion).replace('/', '-')
        self.f_expedicion = str(self.f_expedicion).replace('00:00:00', '').strip()
        if str(self.f_expedicion).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"

        try:
            if self.f_expedicion != '':
                self.f_expedicion = datetime.datetime.strptime(str(self.f_expedicion), format_fecha)
        except Exception as e:
            error = Error()
            error.description = str(
                self.d_excel) + " Se ha producido un error en la fecha " + str(
                self.f_entrega_titulo) + " -- " + e.__str__()
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.normalize_dates"
            # error.log_error()

        if isinstance(self.f_expedicion, datetime.datetime):
            self.f_expedicion = self.f_expedicion.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_expedicion:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_expedicion no tiene formato de fecha REVISAR " + str(self.f_expedicion)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        if isinstance(self.f_recibi_titulo, datetime.datetime):
            self.f_recibi_titulo = self.f_recibi_titulo.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_recibi_titulo:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_recibi_titulo no tiene formato de fecha REVISAR " + str(self.f_recibi_titulo)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        self.f_notif_deneg = str(self.f_notif_deneg).replace('BOE', '')
        self.f_notif_deneg = str(self.f_notif_deneg).replace('None', '')
        self.f_notif_deneg = str(self.f_notif_deneg).replace('/', '-')
        self.f_notif_deneg = str(self.f_notif_deneg).replace('00:00:00', '').strip()
        if str(self.f_notif_deneg).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"
        if self.f_notif_deneg != '':
            self.f_notif_deneg = datetime.datetime.strptime(str(self.f_notif_deneg), format_fecha)

        if isinstance(self.f_notif_deneg, datetime.datetime):
            self.f_notif_deneg = self.f_notif_deneg.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_notif_deneg:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_notif_deneg no tiene formato de fecha REVISAR " + str(self.f_notif_deneg)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        self.f_envio_titulo = str(self.f_envio_titulo).replace('BOE', '')
        self.f_envio_titulo = str(self.f_envio_titulo).replace('None', '')
        self.f_envio_titulo = str(self.f_envio_titulo).replace('/', '-')
        self.f_envio_titulo = str(self.f_envio_titulo).replace('00:00:00', '').strip()
        self.f_envio_titulo = str(self.f_envio_titulo).replace('00-00-0000', '')

        if str(self.f_envio_titulo).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"

        try:
            if self.f_envio_titulo != '':
                self.f_envio_titulo = datetime.datetime.strptime(str(self.f_envio_titulo), format_fecha)
        except Exception as e:
            error = Error()
            error.description = str(
                self.d_excel) + " Se ha producido un error en la fecha " + str(
                self.f_entrega_titulo) + " -- " + e.__str__()
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.normalize_dates"
            # error.log_error()

        if isinstance(self.f_envio_titulo, datetime.datetime):
            self.f_envio_titulo = self.f_envio_titulo.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_envio_titulo:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_envio_titulo no tiene formato de fecha REVISAR " + str(self.f_envio_titulo) + " " + type(self.f_envio_titulo).__name__
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        if isinstance(self.f_desestimiento, datetime.datetime):
            self.f_desestimiento = self.f_desestimiento.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_desestimiento:
                error = Error()
                error.description = str(self.d_excel) + " El campo f_desestimiento no tiene formato de fecha REVISAR " + str(self.f_desestimiento)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()

        self.f_entrega_titulo = str(self.f_entrega_titulo).replace('BOE', '')
        self.f_entrega_titulo = str(self.f_entrega_titulo).replace('None', '')
        self.f_entrega_titulo = str(self.f_entrega_titulo).replace('/', '-')
        self.f_entrega_titulo = str(self.f_entrega_titulo).replace('00:00:00', '').strip()
        self.f_entrega_titulo = str(self.f_entrega_titulo).replace('00-00-0000', '').strip()
        if str(self.f_entrega_titulo).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"

        try:
            if self.f_entrega_titulo != '':
                self.f_entrega_titulo = datetime.datetime.strptime(str(self.f_entrega_titulo), format_fecha)
        except Exception as e:
            error = Error()
            error.description = str(
                self.d_excel) + " Se ha producido un error en la fecha " + str(
                self.f_entrega_titulo) + " -- " + e.__str__()
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.normalize_dates"
            # error.log_error()


        if isinstance(self.f_entrega_titulo, datetime.datetime):
            self.f_entrega_titulo = self.f_entrega_titulo.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_entrega_titulo:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de entrega de titulo no tiene formato de fecha REVISAR " + str(self.f_entrega_titulo)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de envio de titulo ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                # error.log_error() TODO: Dejar como error o no ¿?

        self.fecha_notificacion_boe = str(self.fecha_notificacion_boe).replace('BOE', '')
        self.fecha_notificacion_boe = str(self.fecha_notificacion_boe).replace('None', '')
        self.fecha_notificacion_boe = str(self.fecha_notificacion_boe).replace('/', '-')
        self.fecha_notificacion_boe = str(self.fecha_notificacion_boe).replace('00:00:00', '').strip()
        self.fecha_notificacion_boe = str(self.fecha_notificacion_boe).replace('00-00-0000', '').strip()
        if str(self.fecha_notificacion_boe).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"
        if self.fecha_notificacion_boe != '':
            self.fecha_notificacion_boe = datetime.datetime.strptime(str(self.fecha_notificacion_boe), format_fecha)

        if isinstance(self.fecha_notificacion_boe, datetime.datetime):
            self.fecha_notificacion_boe = self.fecha_notificacion_boe.strftime('%Y-%m-%d')[:10]
        else:
            if self.fecha_notificacion_boe:
                error = Error()
                error.description = str(
                    self.d_excel) + " El campo Fecha de notificacion boe no tiene formato de fecha REVISAR " + str(
                    self.fecha_notificacion_boe)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de notificacion boe ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                # error.log_error() TODO: Dejar como error o no ¿?

        self.f_desestimiento = str(self.f_desestimiento).replace('BOE', '')
        self.f_desestimiento = str(self.f_desestimiento).replace('None', '')
        self.f_desestimiento = str(self.f_desestimiento).replace('/', '-')
        self.f_desestimiento = str(self.f_desestimiento).replace('00:00:00', '').strip()
        self.f_desestimiento = str(self.f_desestimiento).replace('00-00-0000', '').strip()
        if str(self.f_desestimiento).find("-") == 2:
            format_fecha = "%d-%m-%Y"
        else:
            format_fecha = "%Y-%m-%d"
        if self.f_desestimiento != '':
            self.f_desestimiento = datetime.datetime.strptime(str(self.f_desestimiento), format_fecha)

        if isinstance(self.f_desestimiento, datetime.datetime):
            self.f_desestimiento = self.f_desestimiento.strftime('%Y-%m-%d')[:10]
        else:
            if self.f_desestimiento:
                error = Error()
                error.description = str(
                    self.d_excel) + " El campo Fecha de desestimiento no tiene formato de fecha REVISAR " + str(
                    self.fecha_notificacion_boe)
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                error.log_error()
            else:
                error = Error()
                error.description = str(self.d_excel) + " El campo Fecha de desestimiento boe ES NONE "
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.normalize_dates"
                # error.log_error() TODO: Dejar como error o no ¿?

    def set_via_expedicion(self, c_via_exped, comun=0):
        if not c_via_exped:
            self.c_via_exped = '20'
        elif c_via_exped not in (2, 6, 9, 20):
            error = Error()
            if comun == 1:
                error.description = str(self.d_excel) + " Fila: " + str(
                    self.n_fila) + "La via de expedicion no coincide con los valores predefinidos (" + str(
                    c_via_exped) + ") se coloca 20 como valor por defecto"
                self.c_via_exped = 20
            else:
                error.description = str(self.d_excel) + " La via de expedicion no coincide con los valores predefinidos (" + str(c_via_exped) + ") se deja vacia"
                self.c_via_exped = None
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_via_expedicion"
            error.log_error()


        else:
            self.c_via_exped = "0" + str(c_via_exped)

    def set_d_mods(self, d_mods):


        # COMPLETO, PARCIAL, UF, INCOMPLETO O VACÍO
        if not d_mods:
            d_mods = ''
        else:
            d_mods = d_mods.upper().strip()

        if(d_mods in ('COMPLETO', 'PARCIAL', 'IMCOMPLETO', 'INCOMPLETO', 'UNIDAD FORMATIVA', 'UF', '')):
            if self.d_mods == 'IMCOMPLETO':
                self.d_mods = 'INCOMPLETO'

            if self.d_mods == 'UNIDAD FORMATIVA':
                self.d_mods = 'UF'

            self.d_mods = d_mods
        else:
            # No puede ser vacio colocamos completo
            error = Error()
            error.description = str(self.d_excel) + "Modulos formativos no coincide con los valores predefinidos o esta vacio (" + d_mods + ") se coloca COMPLETO por defecto"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_via_expedicion"
            error.log_error()
            self.d_mods = 'COMPLETO'

    def set_c_prop_aprob(self, c_prop_aprob):
        c_prop_aprob = str(c_prop_aprob).strip()
        valores = ('X', 'x')
        if not c_prop_aprob or c_prop_aprob not in valores:
            if c_prop_aprob is not None and c_prop_aprob != 'None':
                error = Error()
                error.description = str(self.d_excel) + " c_prop_aprob (" + str(c_prop_aprob) + ") no coincide con ninguno de los valores se coloca False"
                error.excel_name = self.d_excel
                error.row_number = self.n_fila
                error.function = "CertProf.set_c_prop_aprob"
                error.log_error()
            self.c_prop_aprob = 'N'
        else:
            self.c_prop_aprob = 'P'

    def set_c_prop_denegada(self, c_prop_denegada):
        c_prop_denegada = str(c_prop_denegada).strip()
        if str(c_prop_denegada) == 'None':
            c_prop_denegada = ''
        valores = ('X', 'x', 'X ', 'DENEGADO', 'DENEGAR', 'DESISTIDO ART.71', 'DESISTIDO ART.71 ', 'DESISTIDO ART. 71', 'DESISTIDO ART.91', 'DESISTIDO ART. 91', 'DESISTIDO ART.91 ', 'DESISTIDO ART.68', 'ART. 68', 'DESISTIDO ART. 68', 'DESISTIDO ARTº 68', 'DESISTIDO ART.94', 'DESISTIDO ART. 94', 'DESISTIDO  ART. 94', 'DESISTIMIENTO ART. 94', 'ART 94', 'ACUMULADO', 'ACUMULADO', 'ACUMULADO', 'ACUMULAR', 'ACUMULA', 'SEPE', 'OTRA CA', '')
        if (c_prop_denegada and c_prop_denegada not in valores) and self.c_prop_aprob != 'P':
            error = Error()
            error.description = str(self.d_excel) + " c_prop_denegada (" + c_prop_denegada + ") no coincide con ninguno de los valores se coloca valor vacio por defecto"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.set_c_prop_denegada"
            error.log_error()
            self.c_prop_deneg = ' '
            if c_prop_denegada not in self.arrayDenegada:
                self.arrayDenegada.append(c_prop_denegada)

        else:
            if c_prop_denegada:
                if c_prop_denegada == 'DENEGADO' or c_prop_denegada == 'x' or c_prop_denegada == 'X' or c_prop_denegada == 'X ' or c_prop_denegada == 'DENEGAR':
                    self.c_prop_deneg = 'X'
                elif c_prop_denegada == 'DESISTIDO ART. 94' or c_prop_denegada == 'DESISTIDO  ART. 94' or c_prop_denegada == 'DESISTIMIENTO ART. 94' or c_prop_denegada == 'ART 94' or c_prop_denegada == "DESISTIDO ART.94 ":
                    self.c_prop_deneg = 'DESISTIDO ART.94'
                elif c_prop_denegada == 'DESISTIDO ART. 68' or c_prop_denegada == 'DESISTIDO ART.68' or c_prop_denegada == 'DESISTIDO ART. 68' or c_prop_denegada == 'DESISTIDO ARTº 68' or c_prop_denegada == 'ART 68':
                    self.c_prop_deneg = 'DESISTIDO ART.68'
                elif c_prop_denegada == 'DESISTIDO ART. 71' or c_prop_denegada == 'DESISTIDO ART. 71' or c_prop_denegada == 'DESISTIDO ART.71 ':
                    self.c_prop_deneg = 'DESISTIDO ART.71'
                elif c_prop_denegada == 'DESISTIDO ART.91' or c_prop_denegada == 'DESISTIDO ART. 91':
                    self.c_prop_deneg = 'DESISTIDO ART.91'
                elif c_prop_denegada == 'ACUMULA' or c_prop_denegada == 'ACUMULAR' or 'ACUMULADO' in c_prop_denegada or c_prop_denegada == 'ACUMULADO':
                    self.c_prop_deneg = 'ACUMULADO'

                else:
                    self.c_prop_deneg = c_prop_denegada
            else:
                self.c_prop_deneg = c_prop_denegada

    def existe_persona(self, cursor):
        sql = "sql busca persona'" + self.d_dni[:-1] + "' AND letradni = " + self.d_dni[-1:] + "'"
        cursor.execute(sql)
        row = cursor.fetchone()
        self.per_cod = ''
        if row is None:
            flag = False
        else:
            flag = True
            self.per_fo_cod = row[0]
            self.per_num_doc = row[1]
            self.per_tip_doc = row[2]
            self.per_let_doc = row[3]
        return flag

    def existe_especialidad(self, cursor):

        sql = "sql especialidad'" + self.c_certificado + "'"
        cursor.execute(sql)
        row = cursor.fetchone()
        if row is None:
            flag = False
            error = Error()
            error.description = str(
                self.d_excel) + " No se encuentra la especialidad ( " + str(self.c_certificado) + " ) Sentencia :" +sql
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.existe_especialidad"
            # error.log_error()
        else:
            flag = True
        return flag

    def existe_especialidad_practicas(self, cursor):

        sql = "sql...." + self.c_certificado + "'"
        cursor.execute(sql)
        row = cursor.fetchone()
        if row is None:
            flag = False
            error = Error()
            error.description = str(
                self.d_excel) + " No se encuentra la especialidad de practicas ( " + str(self.c_certificado) + " ) Sentencia :" + sql
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.existe_especialidad_practicas"
            # error.log_error()
        else:
            flag = True
        return flag

    def insertaPersona(self, con, cursor):

        self.per_fo_cod = self.get_per_fo_cod(cursor)
        self.per_num_doc = self.get_per_num_doc(self.d_dni)
        self.per_tip_doc = self.get_per_tip_doc(self.d_dni)
        self.per_let_doc = self.d_dni[-1]


        sql = """sql... 
                VALUES (seq_si_per_fo.nextval, :per_tip_doc, :per_num_doc, :per_let_doc, :per_nom, :per_ape_1, :per_ape_2, TO_DATE(:per_fec_nac,'yyyy-mm-dd'), :pai_cod, :sex_cod)"""
        #Descomentar produccion

        sqlPrint = "sql... VALUES (seq_si_per_fo.nextval, '" + str(self.per_tip_doc) + "', '" + str(self.per_num_doc) + "', '" + str(self.per_let_doc) + "', '" + str(self.d_nombre) + """',q'[""" + str(self.d_apel1) + """]', q'[""" + str(self.d_apel2) + """]',""" + "TO_DATE('" + str(self.f_nacim) + "' ,'yyyy-mm-dd'), 724, " + str(self.c_sexo) + ") RETURNING per_fo_cod INTO v_per_fo_cod ;"
        print(sqlPrint)

        try:
            cursor.execute(sql, (self.per_tip_doc, self.per_num_doc, self.per_let_doc, self.d_nombre, self.d_apel1, self.d_apel2, self.f_nacim, 724, self.c_sexo))
            con.commit()
        except cx_Oracle.IntegrityError as e:
            error = Error()
            error.description = str(self.d_excel) + " No se ha podido guardar la persona ( " + self.d_dni + " ): " + str(e.args[0].code) + " : " + e.args[0].message
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.insertaPersona"
            error.log_error()
        except cx_Oracle.DatabaseError as e:
            error = Error()
            error.description = str(self.d_excel) + " No se ha podido guardar la persona ( " + self.d_dni + " ): " + str(e.args[0].code) + " : " + e.args[0].message
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.insertaPersona"
            error.log_error()

    def insertSiSolCertProf(self, con, logger):
        cursor = con.cursor()

        sql = """sql...
                VALUES(
                :sol_cod, :c_certificado, :per_fo_cod, :per_cod, :per_tip_doc, :per_num_doc, :per_let_doc, TO_DATE(:f_present,'yyyy-mm-dd'), :exp_cod,
                :exp_prv, :exp_ano, :exp_estatal, :c_estado, sysdate, :c_tipo_acredit, :d_observaciones, :d_tecn_prop,
                TO_DATE(:f_prop_resoluc,'yyyy-mm-dd'), TO_DATE(:f_resoluc,'yyyy-mm-dd'), TO_DATE(:f_notif_resoluc,'yyyy-mm-dd'), :c_tipo_resoluc,
                TO_DATE(:f_emit_titulo,'yyyy-mm-dd'), TO_DATE(:f_envio_titulo,'yyyy-mm-dd'), TO_DATE(:f_notif_titulo,'yyyy-mm-dd'), TO_DATE(:f_recep_titulo,'yyyy-mm-dd'), 
                :exp_fam, :dir_nom_via, :dir_mun_cod, :dir_cps_cod, :c_via_exped,
                :c_tipo_solicitud, :c_itinerario, :c_res_informe,
                TO_DATE(:f_notif_boe,'yyyy-mm-dd'), TO_DATE(:f_certificacion,'yyyy-mm-dd'), TO_DATE(:f_desistimiento,'yyyy-mm-dd'),
                :c_registro, TO_DATE(:f_notif_inter,'yyyy-mm-dd'), TO_DATE(:f_entrega_inter,'yyyy-mm-dd'), TO_DATE(:f_result_informe,'yyyy-mm-dd'),
                :check_archivado
                )
        """

        sqlPrintCert = """sql...
                        VALUES(
                        seq_si_sol_cepr.nextval, '""" + str(self.c_certificado) + """', v_per_fo_cod , '""" + str(self.per_cod) + """' , '""" + str(self.per_tip_doc) + """' , '""" + str(self.per_num_doc) + """' , '""" + str(self.per_let_doc) + """' ,  TO_DATE('""" + str(self.f_present) + """','yyyy-mm-dd') , '""" + str(self.exp_cod) + """',
                        '""" + str(self.exp_prv) + """','""" + str(self.exp_ano) + """','""" + str(self.exp_estatal)  + """','""" + str(self.c_estado) + """' , sysdate,'""" + str(self.c_tipo_acreditacion) + """',q'[""" + str(self.d_observaciones) + """]','""" + str(self.d_tecn_prop_reg) + """',
                        TO_DATE('""" + str(self.f_propuesta) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_resolucion) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_notif_deneg) + """','yyyy-mm-dd'), '""" + str(self.c_prop_aprob) + """',
                        TO_DATE('""" + str(self.f_entrega_titulo) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_envio_titulo) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_recibi_titulo) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_recep_titulo) + """','yyyy-mm-dd'), 
                        '""" + str(self.exp_fam) + """','""" + str(self.d_direccion) + """','""" + str(self.d_localidad) + """','""" + str(self.c_postal) + """','""" + str(self.c_via_exped) + """' ,
                        '""" + str(self.c_tipo_acreditacion) + """','""" + str(self.itinerario) + """','""" + str(self.c_res_informe) + """',
                        TO_DATE('""" + str(self.fecha_notificacion_boe) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_certificacion) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_desestimiento) + """','yyyy-mm-dd'),
                        '""" + str(self.c_registro) + """', TO_DATE('""" + str(self.f_recibi_titulo) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_recibi_titulo) + """','yyyy-mm-dd'), TO_DATE('""" + str(self.f_res_informe) + """','yyyy-mm-dd')
                        , '""" + str(self.check_archivado) + """'
                        ) RETURNING SOL_COD INTO v_sol_cod;
                """

        sqlPrintCert = sqlPrintCert.replace("TO_DATE('','yyyy-mm-dd')", "NULL")
        sqlPrintCert = sqlPrintCert.replace("TO_DATE('', 'yyyy-mm-dd')", "NULL")
        sqlPrintCert = sqlPrintCert.replace("TO_DATE(' ', 'yyyy-mm-dd')", "NULL")
        sqlPrintCert = sqlPrintCert.replace("TO_DATE('None','yyyy-mm-dd')", "NULL")
        sqlPrintCert = sqlPrintCert.replace(" 00:00:00", "")
        sqlPrintCert = sqlPrintCert.replace("None", "")

        print(sqlPrintCert)

        if self.get_next_cod_sol(cursor) > 0:
            self.sol_cod = self.get_next_cod_sol(cursor)
        else:
            pass
            # print("Ha ocurrido un error con el sol cod " + self.get_next_cod_sol(cursor))

        try:
            cursor.execute(sql, (
                                 self.sol_cod, self.c_certificado, self.per_fo_cod, self.per_cod, self.per_tip_doc, self.per_num_doc, self.per_let_doc, self.f_present, self.exp_cod,
                                 self.exp_prv, self.exp_ano, self.exp_estatal, self.c_estado, self.c_tipo_acreditacion, self.d_observaciones, self.d_tecn_prop_reg,
                                 self.f_propuesta, self.f_resolucion, self.f_notif_deneg, self.c_prop_aprob,
                                 self.f_entrega_titulo, self.f_envio_titulo, self.f_recibi_titulo, self.f_recep_titulo,
                                 self.exp_fam, self.d_direccion, self.d_localidad, self.c_postal, self.c_via_exped,
                                 self.c_tipo_acreditacion, self.itinerario, self.c_res_informe,
                                 self.fecha_notificacion_boe, self.f_certificacion, self.f_desestimiento,
                                 self.c_registro, self.f_recibi_titulo, self.f_recibi_titulo, self.f_res_informe,
                                 self.check_archivado
                                 ))

            con.commit()

            logger.info(self.d_excel + " se ha guardado con exito")
            return True
        except cx_Oracle.IntegrityError as e:
            error = Error()
            error.description = str(self.d_excel) + " No se ha podido guardar la fila: " + str(e.args[0].code) + " : " + e.args[0].message + str(sqlPrintCert) + str(self.sol_cod)
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.insertSiSolCertProf"
            error.log_error()
            return False
        except cx_Oracle.DatabaseError as e:
            error = Error()
            error.description = str(self.d_excel) + " No se ha podido guardar la fila: " + str(e.args[0].code) + " : " + e.args[0].message
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.insertSiSolCertProf"
            error.log_error()
            return False

    # Metodo para mysql
    def save(self, cursor, connection):
        try:
            sql = """sql... 
                    values(%s, %s ,%s  ,%s ,%s ,%s ,%s ,%s ,%s ,%s 
                    ,%s ,%s ,%s ,%s ,%s ,%s ,%s 
                    ,%s ,%s ,%s ,%s ,%s ,%s ,%s ,%s 
                    ,%s ,%s ,%s ,%s ,%s ,%s 
                    ,%s ,%s ,%s ,%s ,%s ,%s,%s 
                    ,%s ,%s ,%s ,%s ,%s) """

            cursor.execute(sql, (self.d_excel, self.d_pestana, self.n_fila, self.f_present, self.d_dni
                                 , self.d_apel1, self.d_apel2, self.d_nombre, self.c_sexo, self.f_nacim, self.d_direccion
                                 , self.c_postal, self.d_localidad, self.d_provincia, self.c_via_exped, self.c_certificado
                                 , self.d_real_decreto, self.d_certificado, self.f_expedicion
                                 , self.d_fcs, self.d_mods, self.d_doc_present, self.d_observaciones
                                 , self.d_tecn_prop_reg, self.d_doc_req, self.f_doc_req, self.f_doc_pres, self.f_propuesta
                                 , self.c_prop_aprob, self.c_prop_deneg, self.f_resol_deneg, self.f_notif_deneg
                                 , self.c_registro, self.f_resol_prov, self.f_recep_titulo, self.f_envio_titulo
                                 , self.f_entrega_titulo, self.f_recibi_titulo))

            connection.commit()

        except:
            error = Error()
            error.sql = cursor._last_executed
            error.description = "Error desconocido"
            error.excel_name = self.d_excel
            error.row_number = self.n_fila
            error.function = "CertProf.save"
            error.log_error()
            #  print("Error desconocido")

    def get_per_tip_doc(self, dni):
        return "E" if (dni[:1] in ['X', 'Y']) else "D"

    def get_per_num_doc(self, dni):
        dni_f = dni[:-1]
        return dni_f

    def set_c_estado(self):

        if self.d_signatura_arch or self.f_recibi_titulo:
            self.c_estado = '15'  # ARCHIVADO
        elif self.f_recibi_titulo:
            self.c_estado = '12'  # ENTREGADO
        elif self.c_registro:
            self.c_estado = '11'  # EXPEDIDO
        elif self.f_resol_prov:
            self.c_estado = '07'  # ESTIMADO
        elif (self.c_prop_deneg == 'X' or self.c_prop_deneg == 'DENEGADO') and self.f_resol_deneg:
            self.c_estado = '13'  # DENEGADO
        elif (self.c_prop_deneg == 'DESISTIDO ART.71' or self.c_prop_deneg == 'DESISTIDO ART.68') and self.f_resol_deneg:  # Necesita tambien la fecha
            self.c_estado = '06'  # DESESTIMIENTO
        elif (self.c_prop_deneg == 'DESISTIDO ART.91' or self.c_prop_deneg == 'DESISTIDO ART.94' or self.c_prop_deneg == 'DESISTIDO ART. 94') and self.f_resol_deneg:  # Necesita tambien la fecha
            self.c_estado = '05'  # DESESTIMIENTO EXPRESO
        elif self.c_prop_deneg == 'ACUMULADO' or self.c_prop_deneg == 'ACUMULAR':
            self.c_estado = '09'  # ACUMULADO A OTRA SOLICITUD
        elif self.c_prop_deneg == 'SEPE' or self.c_prop_deneg == 'OTRA CA':
            self.c_estado = '10'  # COMP OTRA COM. O SEPE
        elif self.f_propuesta and (self.c_prop_deneg or self.c_prop_aprob):
            self.c_estado = '04'  # INFORMADO
        elif self.f_expedicion or self.c_tipo_acreditacion or self.d_observaciones:
            self.c_estado = '03'  # EN INFORME
        elif self.f_doc_pres or self.d_doc_req != "None" or self.f_doc_req:
            self.c_estado = '02'  # SUBSANACIÓN
        else:
            self.c_estado = '01'

        if self.c_estado == '15':
            self.check_archivado = "S"
        else:
            self.check_archivado = "N"

    def set_expediente_comunes(self):
        arrayRuta = self.d_excel.split('/')
        self.exp_fam = str(arrayRuta[len(arrayRuta) - 1]).replace(".xlsx", "")
        if self.exp_fam == 'COMERCIO':
            self.exp_fam = "COM"  # Solo para COMercio

        self.exp_fam = self.exp_fam.strip()

        self.exp_ano = self.f_present.strftime('%Y')
        self.exp_cod = 2000 + self.n_fila - 7

        prv = str(self.d_provincia).upper()

        if prv == 'AVILA':
            self.exp_prv = '05'
        elif prv == 'BURGOS':
            self.exp_prv = '09'
        elif prv == 'LEON':
            self.exp_prv = '24'
        elif prv == 'PALENCIA':
            self.exp_prv = '34'
        elif prv == 'SALAMANCA':
            self.exp_prv = '37'
        elif prv == 'SEGOVIA':
            self.exp_prv = '40'
        elif prv == 'SORIA':
            self.exp_prv = '42'
        elif prv == 'VALLADOLID':
            self.exp_prv = '47'
        else:
            self.exp_prv = '47'

    def set_tipo_acreditacion(self):
        opciones = {'COMPLETO': 'T', 'PARCIAL': 'P', 'INCOMPLETO': 'I', 'ANOTACION': 'A'}
        if self.d_mods not in opciones or not self.d_mods:
            self.c_tipo_acreditacion = 'T'
        elif self.d_mods and self.d_mods in opciones:
            self.c_tipo_acreditacion = opciones[self.d_mods]

    def set_f_notificacion_boe(self):
        if 'boe' in str(self.f_notif_deneg) or 'BOE' in str(self.f_notif_deneg):
            if self.f_notif_deneg:
                self.fecha_notificacion_boe = str(self.f_notif_deneg).replace("BOE", "")
                self.fecha_notificacion_boe = str(self.f_notif_deneg).replace("boe", "")
        else:
            self.fecha_notificacion_boe = ''

    def set_f_desestimiento(self):
        # Si En AE denegado X la fecha es col AC fecha propuesta tecnico
        if self.c_prop_deneg == 'X':
            self.f_desestimiento = self.f_propuesta
        else:
            self.f_desestimiento = None

    def set_f_envio_otra_ca(self):
        self.f_desestimiento = None


    def normalize_doc(self, doc):
        doc = str(doc).replace('-', ',')
        doc = str(doc).replace('/', ',')
        doc = str(doc).replace(';', ',')
        doc = str(doc).replace('Y', ',')
        doc = str(doc).replace('y', ',')

        return doc

    def validarDNI(self, dni):
        tabla = "TRWAGMYFPDXBNJZSQVHLCKE"
        dig_ext = "XYZ"
        reemp_dig_ext = {'X': '0', 'Y': '1', 'Z': '2'}
        numeros = "1234567890"
        dni = dni.upper()
        if len(dni) == 9:
            dig_control = dni[8]
            dni = dni[:8]
            if dni[0] in dig_ext:
                dni = dni.replace(dni[0], reemp_dig_ext[dni[0]])
            return len(dni) == len([n for n in dni if n in numeros]) \
                   and tabla[int(dni) % 23] == dig_control
        return False



