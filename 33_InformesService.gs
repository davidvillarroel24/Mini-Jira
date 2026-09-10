function _buildApiResponse(ok, data, error){

	return {
		ok: ok,
		data: data || null,
		error: error || null
	};
}

function _mapInformeError(error){

	const message = error && error.message ? String(error.message) : "Error interno";

	if(message.indexOf("AUTH_REQUIRED") === 0){
		return {
			code: "AUTH_REQUIRED",
			message: "Debe iniciar sesion"
		};
	}

	if(message.indexOf("VALIDATION_ERROR") === 0){
		return {
			code: "VALIDATION_ERROR",
			message: message.replace("VALIDATION_ERROR:", "").trim() || "Solicitud invalida"
		};
	}

	return {
		code: "INTERNAL_ERROR",
		message: message,
		details: error && error.stack ? String(error.stack) : ""
	};
}

function previewExportReport(payload){

	try{

		const htmlPreview = generarPreviewInforme(payload, "preview");

		const reporte = _buildInformeDesarrollo(payload, "preview");

		return _buildApiResponse(true, {
			htmlPreview: htmlPreview,
			resumen: reporte.resumen,
			imageDiagnostics: reporte.imageDiagnostics || []
		}, null);

	}catch(error){

		return _buildApiResponse(false, null, _mapInformeError(error));
	}
}

function exportReportPdf(payload){

	try{

		const pdf = generarPDF(payload);
		const reporte = _buildInformeDesarrollo(payload, "pdf", payload && payload.__user);

		return _buildApiResponse(true, {
			base64: pdf.base64,
			nombreArchivo: pdf.nombre,
			imageDiagnostics: (reporte && reporte.imageDiagnostics) || []
		}, null);

	}catch(error){

		return _buildApiResponse(false, null, _mapInformeError(error));
	}
}
