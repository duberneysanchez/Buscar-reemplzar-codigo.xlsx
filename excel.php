<?php

// Incluir la librería PhpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
ini_set('memory_limit', '512M'); // Ajusta el límite de memoria según tus necesidades

// Carpeta de los archivos Excel
$directorio = 'C:/Users/USERS/Desktop/excelphp/';

// Listado xlsx que contiene la información
$rutaListado = 'C:/Users/USERS/Desktop/cambio.xlsx';

// Cargar el archivo listado.xlsx
$listado = IOFactory::load($rutaListado);
$hojaListado = $listado->getActiveSheet();

// Array para almacenar los códigos anteriores y nuevos
$cambios = [];

//  Códigos a buscar y reemplazar
foreach ($hojaListado->getRowIterator() as $fila) {
    // Número de fila
    $numeroFila = $fila->getRowIndex();
   
    // Celdas a buscar en las columnas A y B del listado
    $codigoAnterior = $hojaListado->getCell('B'.$numeroFila)->getValue();
    $codigoNuevo = $hojaListado->getCell('A'.$numeroFila)->getValue();
   //getCalculatedValue usa si la celda tiene  fórmula
    // Almacenar los códigos anteriores y nuevos en el array $cambios
    $cambios[$codigoAnterior] = $codigoNuevo;
}

// Lista de archivos en el directorio
$archivos = scandir($directorio);
echo "Iniciando procesamiento...\n";

// Recorrer cada archivo en el directorio
foreach ($archivos as $archivo) {
    // Si el archivo es un Excel (.xlsx)
    if (pathinfo($archivo, PATHINFO_EXTENSION) == 'xlsx') {
        echo "Procesando archivo: $archivo\n";
        // Cargar el archivo Excel
        $objPHPExcel = IOFactory::load($directorio . $archivo);
        $hojaActual = $objPHPExcel->getActiveSheet();
		
		$reemplazoRealizado = false; // Variable para rastrear si se realizó al menos un reemplazo en este archivo

        // Recorrer solo el rango de celdas deseado (por ejemplo, A1:A1000)
        foreach ($hojaActual->getRowIterator(2, 100) as $row) {
            foreach ($row->getCellIterator('A', 'B') as $celda) {
                // Valor de cada celda
                $valorCelda = $celda->getValue();

                // Verificar si $valorCelda no es null antes de llamar a trim()
                if ($valorCelda !== null) {
                    $valorCelda = trim($valorCelda);
                } else {
                    // Asignar un valor por defecto o tomar alguna acción según sea necesario
                    $valorCelda = ''; // O cualquier otro valor predeterminado que desees
                }

                // Este código es para buscar y reemplazar el código principal sin la versión adicional
                foreach ($cambios as $codigoAnterior => $codigoNuevo) {
                    // Verificar si el código anterior está presente en el valor de la celda
                    if (strpos($valorCelda, $codigoAnterior) !== false) {
                        // Obtener la dirección de la celda
                        $direccionCelda = $celda->getCoordinate();

                        // Verificar si $codigoNuevo no es null antes de usarlo
                        if ($codigoNuevo !== null) {
                            // Reemplazar solo el código anterior manteniendo la versión existente
                            $valorNuevo = str_replace($codigoAnterior, $codigoNuevo, $valorCelda);

                            // Actualizar el valor de la celda con el código actualizado
                            $hojaActual->setCellValue($direccionCelda, $valorNuevo);
                        }

                        // Salir del bucle una vez que se encuentra una coincidencia para evitar múltiples reemplazos en una celda
						$reemplazoRealizado = true;
                        break;
                    }
                }
                //echo "Valor de la celda: $valorCelda\n";
                // Verificar si la cadena comienza con "Vigente desde:"
                if (strpos($valorCelda, "Vigente desde:") === 0) {
                    // Extraer la fecha del formato YYYY-MM-DD
                    preg_match('/Vigente desde:\s*(\d{4}-\d{2}-\d{2})/', $valorCelda, $matches);
                    $fechaActual = isset($matches[1]) ? $matches[1] : null;

                    // Reemplazar la fecha actual con "2024-03-01"
                    if ($fechaActual !== null) {
                        $valorCelda = str_replace($fechaActual, '2024-03-01', $valorCelda);
                    }
                    echo "Celda actualizada: $valorCelda\n";
                    // Obtener la dirección de la celda
                    $direccionCelda = $celda->getCoordinate();

                    // Actualizar el valor de la celda con la cadena modificada
                    $hojaActual->setCellValue($direccionCelda, $valorCelda);
                }
            }
        }

        // Guardar los cambios en el archivo Excel
        if ($reemplazoRealizado) {
		$writer = IOFactory::createWriter($objPHPExcel, 'Xlsx');
        $writer->save($directorio . $archivo);
		}
    }
}

echo "Proceso completado.";
