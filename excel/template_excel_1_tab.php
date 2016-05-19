<?PHP

/* ==================================================================== <HEADER>
 * Title       : **NAAMBESTAND**
 * Description :
 * =========================================================== <PROGRAM HISTORY>
 * Wijzigingen in git zie http://dwh.mchaaglanden.local/gitphp/?sort=age
 * ===================================================================== <NOTES>
 *
 * ==================================================================== <SOURCE>
 */

/*
 * Variabelen aan te passen voor de werking van het bestand
 */

    $naamtab = '**NAAMTAB**';
    $naambestand = '**NAAMBESTAND**';
    $pad = '//mchaaglanden/mchdfs/DWH/Algemeen/bestanden/uit/';
    $onderwerp = 'Maandelijks bestand';
    // Ontvangers(datawarehous krijgt standaard een kopie)
    $to = array('datawarehouse@mchaaglanden.nl','**ONTVANGER**');
    $message  = '
        Beste collega,

        Bijgevoegd een maandelijks bestand.

        --
        Datawarehouse Medisch Centrum Haaglanden
        Onze producten http://dwh.mchaaglanden.local/index/

        '.$pad.$filename.'
        ';

    // Periode die in de naam van het bestand moet komen, standaard jaarmaanddag
    $now = date('Ymd');
    // Bestand verwijderen na mailen?
    $verwijderenverwijderen = true;

/*
 * Optioneel: geef aan welke velden "echt" tekst zijn(voorloopnullen)
 * We beginnen met tellen bij 0
 */

    $arrTEXTColumns = array();
    // **OPTIONEEL AANPASSEN**
    // We teller vanaf kolom 0. Zet een kolom om 1 om er een tekstveld van te maken
    // $arrTEXTColumns[0] = 1;
    // $arrTEXTColumns[2] = 1;

function DisplayErrors()
{
    $errors = sqlsrv_errors(SQLSRV_ERR_ERRORS);
    foreach ($errors as $error) {
          echo "Error: ".$error['message']."\n";
    }
}

include('connect_nt-vm-dwh-p3.php');
include('htmlMimeMail5.php');
include('PHPExcel.php');
include('PHPExcel/Writer/Excel2007.php');

echo "bezig met tablad 1 (van 1) ..." ."\r\n";
$sql = GetSQL1();

@set_time_limit(0);

echo "Start SQL ..." ."\r\n";

$result = sqlsrv_query($DWH_EAIB, $sql, array(), array('Scrollable' => 'static'));

if ($result === false) {
    DisplayErrors();
    die($sql);
}

$arrFields   = sqlsrv_field_metadata($result);
$nrOfFields  = sqlsrv_num_fields($result);

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
// $naamtab is aan te passen bij de variabelen
$objPHPExcel->getActiveSheet()->setTitle($naamtab);

$rowIndex = 1;

for ($i = 0; $i <$nrOfFields; ++$i) {
    $colString = PHPExcel_Cell::stringFromColumnIndex($i);
    $objPHPExcel->getActiveSheet()->SetCellValue($colString.$rowIndex, $arrFields[$i]['Name']);
}

$rowIndex++;

@set_time_limit(0);

while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_NUMERIC)) {

    for ($i = 0; $i < sizeof($row); ++$i) {
        $colString = PHPExcel_Cell::stringFromColumnIndex($i);

        if (isset($arrTEXTColumns[$i])) {
            $objPHPExcel->getActiveSheet()->getCell($colString.$rowIndex)->setValueExplicit($row[$i], PHPExcel_Cell_DataType::TYPE_STRING);
        } else {
            $objPHPExcel->getActiveSheet()->SetCellValue($colString.$rowIndex, $row[$i]);
        }
    }

    $rowIndex++;
}

//Even zeker maken dat er niets meer in de buffer is blijven hangen.
while (@ob_end_clean());

@set_time_limit(0);

//Excel maken en verzenden

echo "Mailen van de excel ..." ."\r\n";
$objPHPExcel->setActiveSheetIndex(0);

$filename = $naambestand.$now.'.xlsx';

$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save($pad.$filename);

$from     = 'datawarehouse@mchaaglanden.nl';
$subject  = $onderwerp;

$mail = new htmlMimeMail5();
$mail->setFrom($from);
$mail->setSubject($subject);
$mail->setText($message);
$mail->addAttachment(new fileAttachment($pad.$filename));
$mail->send($to);

if ($verwijderen) {
    unlink($pad.$filename);
}


return;

function GetSQL1()
{
    return
    "

set nocount on -- Stop de melding over aantal regels
set ansi_warnings on -- ISO foutmeldingen(NULL in aggregraat bv)
set ansi_nulls on -- ISO NULLL gedrag(field = null returns null, ook als field null is)

**SQL**

    ";

}
