function crearPresupuesto(presupuestoNumero, presupuestoCliente, presupuestoAnunciante, presupuestoFecha, presupuestoCostoColocacion, presupuestoTotal, presupuestoDescuento, presupuestoTotalConDescuento, presupuestoContacto, presupuestoPropiedades, presupuestoUsuario, presupuestoTituloInformacionAdicional, presupuestoInformacionAdicional, presupuestoTituloDescuento, presupuestoTituloTotalConDescuento, presupuestoCondiciones, presupuestoVersion) {

    presupuestoNumero = presupuestoNumero === null ? '' : presupuestoNumero;
    presupuestoCliente = presupuestoCliente === null ? '' : presupuestoCliente.toUpperCase();
    presupuestoAnunciante = presupuestoAnunciante === null ? '' : presupuestoAnunciante.toUpperCase();
    presupuestoFecha = presupuestoFecha === null ? '' : presupuestoFecha;
    presupuestoCostoColocacion = presupuestoCostoColocacion === null ? '' : `$${Number(presupuestoCostoColocacion).toLocaleString('es-ES')}`;
    presupuestoTotal = presupuestoTotal === null ? '' : `$${Number(presupuestoTotal).toLocaleString('es-ES')}`;
    presupuestoDescuento = presupuestoDescuento === null ? '' : `$${Number(presupuestoDescuento).toLocaleString('es-ES')}`;
    presupuestoTotalConDescuento = presupuestoTotalConDescuento === null ? '' : `$${Number(presupuestoTotalConDescuento).toLocaleString('es-ES')}`;
    presupuestoContacto = presupuestoContacto === null ? '' : presupuestoContacto.toUpperCase();
    presupuestoPropiedades = presupuestoPropiedades === null ? '' : presupuestoPropiedades;
    presupuestoUsuario = presupuestoUsuario === null ? '' : presupuestoUsuario.toUpperCase();
    presupuestoCondiciones = presupuestoCondiciones === null ? '' : presupuestoCondiciones;
    presupuestoVersion = presupuestoVersion === null ? '' : presupuestoVersion;

    const spreadsheet = SpreadsheetApp.openById("1ZWmSpmfCoNMy_CEAOZAaj-I5JDLGYxEL__D1hGU4wiE");
    const sheetInformacion = spreadsheet.getSheetByName("Propiedad");
    const sheetPropiedades = spreadsheet.getSheetByName("PresupuestoPropiedad");

    const datosInformacion = convertirDatosAObjetos(sheetInformacion);
    const datosPropiedades = convertirDatosAObjetos(sheetPropiedades);


    const presupuestosFolder = DriveApp.getFolderById('1E1QLDI6byjhpR-DWK67-41liTM2r_c68');

    const presupuestoTemplateId = '1jopQ5yQaMfq9Y2M-dobu4evMuyiMBkq9vPB_bpxsSVU';


    const formattedPresupuestoFecha = formatDate(presupuestoFecha);

    const presupuestoVersionAnterior = presupuestoVersion - 1;

    const presupuestosExistentes = presupuestosFolder.getFilesByName('Presupuesto ' + presupuestoNumero + ' v' + presupuestoVersionAnterior + '.pdf');
    while (presupuestosExistentes.hasNext()) {
        const presupuestoActual = presupuestosExistentes.next();
        const presupuestoEliminar = presupuestoActual.getId();
        const presupuestoExiste = DriveApp.getFileById(presupuestoEliminar);
        if (presupuestoExiste) {
            presupuestoExiste.setTrashed(true);
        }
    }

    const presupuestoId = DriveApp.getFileById(presupuestoTemplateId).makeCopy(presupuestosFolder).getId();
    DriveApp.getFileById(presupuestoId).setName('Presupuesto ' + presupuestoNumero + ' v' + presupuestoVersion);

    const presupuestoDoc = DocumentApp.openById(presupuestoId);
    const presupuestoHeader = presupuestoDoc.getHeader();
    const presupuestoBody = presupuestoDoc.getBody();
    const presupuestoFooter = presupuestoDoc.getFooter();

    presupuestoBody.replaceText('##CLIENTE##', presupuestoCliente);
    presupuestoBody.replaceText('##FECHA##', formattedPresupuestoFecha);
    presupuestoBody.replaceText('##ANUNCIANTE##', presupuestoAnunciante);
    presupuestoBody.replaceText('##CONTACTO##', presupuestoContacto);
    presupuestoBody.replaceText('##COSTOCOLOCACION##', presupuestoCostoColocacion);
    presupuestoBody.replaceText('##TOTALPRESUPUESTO##', presupuestoTotal + ".- + IVA");
    presupuestoBody.replaceText('##USUARIO##', presupuestoUsuario);
    presupuestoBody.replaceText('##TITULOINFORMACIONADICIONAL##', presupuestoTituloInformacionAdicional);
    presupuestoBody.replaceText('##INFORMACIONADICIONAL##', presupuestoInformacionAdicional);
    presupuestoBody.replaceText('##CONDICIONES##', presupuestoCondiciones);


    if (presupuestoDescuento === "$" + 0) {

        const elementosAEliminar = [
            presupuestoBody.findText('##TITULODESCUENTO##').getElement().getParent().getParent(),
            presupuestoBody.findText('##TITULOTOTALCONDESCUENTO##').getElement().getParent().getParent(),
            presupuestoBody.findText('##DESCUENTO##').getElement().getParent().getParent(),
            presupuestoBody.findText('##TOTALCONDESCUENTO##').getElement().getParent().getParent()
        ];
        elementosAEliminar.forEach(elemento => elemento.removeFromParent());

    } else {

        presupuestoBody.replaceText('##DESCUENTO##', presupuestoDescuento + ".-");
        presupuestoBody.replaceText('##TOTALCONDESCUENTO##', presupuestoTotalConDescuento + ".- + IVA");
        presupuestoBody.replaceText('##TITULODESCUENTO##', presupuestoTituloDescuento);
        presupuestoBody.replaceText('##TITULOTOTALCONDESCUENTO##', presupuestoTituloTotalConDescuento);
    }

    const tabla = presupuestoDoc.getTables()[1];

    presupuestoPropiedades.forEach(id => {
        const rowPropiedadEncontrada = buscarInformacionPorID(id, datosPropiedades);
        if (rowPropiedadEncontrada) {

            const valorExhibicion = `$${Number(rowPropiedadEncontrada.pp_precio_mensual).toLocaleString('es-ES')}`;

            const rowInformacion = datosInformacion.find(e => e.propiedad_id === rowPropiedadEncontrada.pp_propiedad);


            const pdfURL = rowInformacion.propiedad_pdf;
            const mapURL = rowInformacion.propiedad_link_geolocalizacion;
            const ref = rowInformacion.propiedad_ref_n;
            const ubicacion = rowInformacion.propiedad_ubicacion;
            const localidad = rowInformacion.propiedad_localidad;
            const latlong = rowInformacion.propiedad_geolocalizacion;
            const tipo = rowInformacion.propiedad_tipo;
            const caras = rowInformacion.propiedad_caras;
            const base = rowInformacion.propiedad_base;
            const alto = rowInformacion.propiedad_altura;

            const filaExistente = tabla.getRow(1).copy();
            const nuevaFila = filaExistente.copy();
            const refElement = nuevaFila.findText('##REF##').getElement();
            const latlongElement = nuevaFila.findText('##LATLONG##').getElement();

            refElement.asText().setText(ref).setLinkUrl(pdfURL);
            refElement.setFontSize(10);

            latlongElement.asText().setText(latlong).setLinkUrl(mapURL);
            latlongElement.setFontSize(10);

            nuevaFila.replaceText('##VALOR##', valorExhibicion);
            nuevaFila.replaceText('##UBICACION##', ubicacion);
            nuevaFila.replaceText('##LOCALIDAD##', localidad);
            nuevaFila.replaceText('##TIPO##', tipo);
            nuevaFila.replaceText('##CARAS##', caras);
            nuevaFila.replaceText('##BASE##', base);
            nuevaFila.replaceText('##ALTO##', alto);

            tabla.appendTableRow(nuevaFila);
        }

    });

    tabla.removeRow(1);

    presupuestoDoc.saveAndClose();

    const pdfFile = DriveApp.getFileById(presupuestoId).getAs('application/pdf');
    DriveApp.getFolderById('1E1QLDI6byjhpR-DWK67-41liTM2r_c68').createFile(pdfFile);

    DriveApp.getFileById(presupuestoId).setTrashed(true);
}




function formatDate(dateString) {
    const [year, month, day] = dateString.split('T')[0].split('-');
    return `${day}/${month}/${year}`;
}

function convertirDatosAObjetos(sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    const objectsArray = data.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });

    return objectsArray;
}

function buscarInformacionPorID(id, datosPropiedades) {
    return datosPropiedades.find(propiedad => propiedad.pp_id === id);
}
