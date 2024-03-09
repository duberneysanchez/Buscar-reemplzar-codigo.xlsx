<?php

// librería PhpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// Carpeta de los archivos Excel
$directorio = 'C:/Users/USERS/Desktop/excelphp/';

// Listado xlsx que contiene la información
$rutaListado = 'C:/Users/USERS/Desktop/listado.xlsx';

// variable códigos de reemplazo
$listado = IOFactory::load($rutaListado);
$hojaListado = $listado->getActiveSheet();

// array para almacenar los códigos anteriores y nuevos
$cambios = [];

//  códigos a buscar y reemplazar
foreach ($hojaListado->getRowIterator() as $fila) {
    // número de fila
    $numeroFila = $fila->getRowIndex();
   
    // celdas a buscar A y B del listado
    $codigoAnterior = $hojaListado->getCell('B'.$numeroFila)->getValue();
    $codigoNuevo = $hojaListado->getCell('D'.$numeroFila)->getValue();
   
    // array códigos anteriores-nuevos
    $cambios[$codigoAnterior] = $codigoNuevo;
}


// lista de archivos en el directorio
$archivos = scandir($directorio);

//  Recorremos cada archivo en el directorio
foreach ($archivos as $archivo) {
    // archivo Excel .xlsx
    if (pathinfo($archivo, PATHINFO_EXTENSION) == 'xlsx') {
        // Cargamos el archivo Excel
        $objPHPExcel = IOFactory::load($directorio . $archivo);
        $hojaActual = $objPHPExcel->getActiveSheet();

        // revisar cada celda
        foreach ($hojaActual->getRowIterator() as $row) {
            foreach ($row->getCellIterator() as $celda) {
                // valor de cada celda
                $valorCelda = $celda->getValue();

                // Verificamos si $valorCelda no es null antes de llamar a trim()
                if ($valorCelda !== null) {
                    $valorCelda = trim($valorCelda);
                } else {
                    // Asignamos un valor por defecto o tomamos alguna acción según sea necesario
                    $valorCelda = ''; // O cualquier otro valor predeterminado que desees
                }

                // Este código es para buscar y reemplazar el código principal sin la versión adicional
                foreach ($cambios as $codigoAnterior => $codigoNuevo) {
                    // Verificamos si el código anterior está presente en el valor de la celda
                    if (strpos($valorCelda, $codigoAnterior) !== false) {
                        // Obtenemos la dirección de la celda
                        $direccionCelda = $celda->getCoordinate();

                        // Verificamos si $codigoNuevo no es null antes de usarlo
                        if ($codigoNuevo !== null) {
                            // Reemplazamos solo el código anterior manteniendo la versión existente
                            $valorNuevo = str_replace($codigoAnterior, $codigoNuevo, $valorCelda);

                            // Actualizamos el valor de la celda con el código actualizado
                            $hojaActual->setCellValue($direccionCelda, $valorNuevo);
                        }

                        // Salimos del bucle una vez que encontramos una coincidencia para evitar múltiples reemplazos en una celda
                        break;
                    }
                }

               // Verificamos si la cadena comienza con "Vigente desde:"
if (strpos($valorCelda, "Vigente desde:") === 0) {
    // Extraemos la fecha del formato YYYY-MM-DD
    preg_match('/Vigente desde:\s*(\d{4}-\d{2}-\d{2})/', $valorCelda, $matches);
    $fechaActual = isset($matches[1]) ? $matches[1] : null;

    // Reemplazamos la fecha actual con "2024-03-01"
    if ($fechaActual !== null) {
        $valorCelda = str_replace($fechaActual, '2024-03-01', $valorCelda);
    }

    // Obtenemos la dirección de la celda
    $direccionCelda = $celda->getCoordinate();

    // Actualizamos el valor de la celda con la cadena modificada
    $hojaActual->setCellValue($direccionCelda, $valorCelda);
                }
            }
        }

        // Guardar los cambios en el archivo Excel
        $writer = IOFactory::createWriter($objPHPExcel, 'Xlsx');
        $writer->save($directorio . $archivo);
    }
}

echo "Proceso completado.";
