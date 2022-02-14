
class BBDD(object):
        """ Sentencias sql """
        create_table = """
            create table if not exists cert_prof_xls_import(
            d_excel varchar(100) not null,
            d_pestana varchar(20) not null,
            n_fila int(3) not null,
            f_present date,
            d_dni varchar(9),
            d_apel1 varchar(40),
            d_apel2 varchar(40),
            d_nombre varchar(40),
            c_sexo char(1),
            f_nacim date,
            d_direccion varchar(100),
            c_postal varchar(5),
            d_localidad varchar(50),
            d_provincia varchar(40),
            c_via_exped int(9),
            c_certificado varchar(10),
            c_nivel int(5),
            d_real_decreto varchar(100),
            d_certificado varchar(100),
            x_horas int(5),
            f_expedicion date,
            d_fcs varchar(100),
            d_mods varchar(40),
            d_doc_present varchar(40),
            d_centro_exped varchar(20),
            d_observaciones varchar(512),
            d_tecn_prop_reg varchar(10),
            d_doc_req varchar(100),
            f_doc_req date,
            f_doc_pres date,
            f_propuesta date,
            c_prop_aprob char(1),
            c_prop_deneg char(1),
            f_resol_deneg date,
            f_notif_deneg date,
            f_silcoi date,
            c_registro varchar(11),
            f_resol_prov date,
            f_listado_silcoi date,
            f_recep_titulo date,
            f_envio_titulo date,
            f_entrega_titulo date,
            f_recibi_titulo date
        )"""
        clean_db = "delete from cert_prof_xls_import "
                
        pass
