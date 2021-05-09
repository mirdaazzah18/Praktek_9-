<?php
//membuka koneksi ke database
include "koneksi1.php";
//memanggil library
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//menuliskan nama kolom pada excel
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Jenis Pendaftaran');
$sheet->setCellValue('B1', 'Tanggal Masuk');
$sheet->setCellValue('C1', 'NIS');
$sheet->setCellValue('D1', 'Nomor Peserta Ujian');
$sheet->setCellValue('E1', 'Pernah Paud?');
$sheet->setCellValue('F1', 'Pernah TK?');
$sheet->setCellValue('G1', 'No SKHUN Sebelumnya');
$sheet->setCellValue('H1', 'No Ijazah Sebelumnya');
$sheet->setCellValue('I1', 'Hobi');
$sheet->setCellValue('J1', 'Cita-Cita');
$sheet->setCellValue('K1', 'Nama Lengkap');
$sheet->setCellValue('L1', 'Jenis Kelamin');
$sheet->setCellValue('M1', 'No NISN');
$sheet->setCellValue('N1', 'No NIK');
$sheet->setCellValue('O1', 'Tempat Lahir');
$sheet->setCellValue('P1', 'Tanggal Lahir');
$sheet->setCellValue('Q1', 'Agama');
$sheet->setCellValue('R1', 'Berkebutuhan Khusus');
$sheet->setCellValue('S1', 'Alamat');
$sheet->setCellValue('T1', 'RT');
$sheet->setCellValue('U1', 'RW');
$sheet->setCellValue('V1', 'Nama Dusun');
$sheet->setCellValue('W1', 'Nama Kelurahan/Desa');
$sheet->setCellValue('X1', 'Nama Kecamatan');
$sheet->setCellValue('Y1', 'Kode Pos');
$sheet->setCellValue('Z1', 'Tempat Tinggal');
$sheet->setCellValue('AA1', 'Moda Transportasi');
$sheet->setCellValue('AB1', 'No HP');
$sheet->setCellValue('AC1', 'No Telp');
$sheet->setCellValue('AD1', 'Email Pribadi');
$sheet->setCellValue('AE1', 'Penerima KPS/PKH/KIP');
$sheet->setCellValue('AF1', 'No KPS/PKH/KIP');
$sheet->setCellValue('AG1', 'Kewarganegaraan');

//mengambil data pada database dan menuliskan pada excel
$query = mysqli_query($conn,"select * from pendaftaran");
$i = 2;
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $row['JENIS_PENDAFTARAN']);
	$sheet->setCellValue('B'.$i, $row['TANGGAL_MASUK']);
	$sheet->setCellValue('C'.$i, $row['NIS']);
	$sheet->setCellValue('D'.$i, $row['NOMOR_PESERTA']);
	$sheet->setCellValue('E'.$i, $row['PAUD']);
	$sheet->setCellValue('F'.$i, $row['TK']);
	$sheet->setCellValue('G'.$i, $row['NO_SKHUN']);
	$sheet->setCellValue('H'.$i, $row['NO_IJAZAH']);
	$sheet->setCellValue('I'.$i, $row['HOBI']);
	$sheet->setCellValue('J'.$i, $row['CITA_CITA']);
	$sheet->setCellValue('K'.$i, $row['JENIS_KELAMIN']);
	$sheet->setCellValue('L'.$i, $row['NAMA']);
	$sheet->setCellValue('M'.$i, $row['NISN']);
	$sheet->setCellValue('N'.$i, $row['NIK']);
	$sheet->setCellValue('O'.$i, $row['TEMPAT_LAHIR']);
	$sheet->setCellValue('P'.$i, $row['TANGGAL_LAHIR']);
	$sheet->setCellValue('Q'.$i, $row['AGAMA']);
	$sheet->setCellValue('R'.$i, $row['BERKEBUTUHAN_KHUSUS']);
	$sheet->setCellValue('S'.$i, $row['ALAMAT']);
	$sheet->setCellValue('T'.$i, $row['RT']);
	$sheet->setCellValue('U'.$i, $row['RW']);
	$sheet->setCellValue('V'.$i, $row['DUSUN']);
	$sheet->setCellValue('W'.$i, $row['KELURAHAN']);
	$sheet->setCellValue('X'.$i, $row['KECAMATAN']);
	$sheet->setCellValue('Y'.$i, $row['KODE_POS']);
	$sheet->setCellValue('Z'.$i, $row['TEMPAT_TINGGAL']);
	$sheet->setCellValue('AA'.$i, $row['TRANSPORTASI']);
	$sheet->setCellValue('AB'.$i, $row['NO_HP']);
	$sheet->setCellValue('AC'.$i, $row['NO_TELP']);
	$sheet->setCellValue('AD'.$i, $row['EMAIL']);
	$sheet->setCellValue('AE'.$i, $row['PENERIMA_KPS']);
	$sheet->setCellValue('AF'.$i, $row['NO_KPS']);
	$sheet->setCellValue('AG'.$i, $row['KEWARGANEGARAAN']);		
	$i++;
}

//style
$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:Y'.$i)->applyFromArray($styleArray);

//memunculkan file excel
$writer = new Xlsx($spreadsheet);
$writer->save('Report Pendaftaran Siswa Baru.xlsx');
?>