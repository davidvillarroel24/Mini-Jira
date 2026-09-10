function generarPDF(idsOrPayload){

  // =========================
  // HTML
  // =========================
    
  const html =
    generarPreviewInforme(
      idsOrPayload,
      "pdf"
    );

  // =========================
  // PDF
  // =========================

  const blob =

    HtmlService

      .createHtmlOutput(html)

      .getBlob()

      .getAs('application/pdf');

  // =========================
  // FECHA
  // =========================

  const fecha =

    Utilities.formatDate(

      new Date(),

      Session.getScriptTimeZone(),

      "yyyyMMdd_HHmm"
    );

  // =========================
  // NOMBRE
  // =========================

  const nombre =
    `Informe_Desarrollo_${fecha}.pdf`;

  // =========================
  // BASE64
  // =========================

  const base64 =

    Utilities.base64Encode(
      blob.getBytes()
    );

  // =========================
  // RETURN
  // =========================

  return {

    base64:
      base64,

    nombre:
      nombre
  };
}

function obtenerDataPDF(ids){

  return DB_OBS_getAll()

    .filter(x =>

      ids.includes(
        x.ID_OBSERVACION
      )
    )

    .map(obs => {

      let imageBase64 = "";

      // =========================
      // IMAGEN
      // =========================

      try{

        if(obs.CAPTURAS){

          const match =

            obs.CAPTURAS.match(
              /[-\w]{25,}/
            );

          if(match){

            const fileId =
              match[0];

            const file =

              DriveApp
                .getFileById(
                  fileId
                );

            // =========================
            // JPG
            // =========================

            const blob =

              file

                .getBlob()

                .getAs(
                  MimeType.JPEG
                );

            imageBase64 =

              Utilities.base64Encode(
                blob.getBytes()
              );
          }
        }

      }catch(err){

        Logger.log(err);
      }

      // =========================
      // RETURN
      // =========================

      return {

        ID_OBSERVACION:
          obs.ID_OBSERVACION,

        FECHA:
          obs.FECHA,

        QA:
          obs.QA,

        MODULO:
          obs.MODULO,

        ESTADO:
          obs.ESTADO,

        PRIORIDAD:
          obs.PRIORIDAD,

        DESCRIPCION_CORTA:
          obs.DESCRIPCION_CORTA,

        PASOS_REPRODUCIR:
          obs.PASOS_REPRODUCIR,

        IMAGEN_BASE64:
          imageBase64
      };
    });
}
