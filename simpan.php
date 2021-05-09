<?php
//mengambil data dari formpendaftaran.php
include 'formpendaftaran.php';

//penampilan data kembali
echo "<h2> Your Input</h2>";
echo "jenis_pendaftaran = ".$jenis_pendaftaran;
echo "<br>";
echo "tanggal_masuk = ".$tanggal_masuk;
echo "<br>";
echo "nis = ".$nis;
echo "<br>";
echo "nomor_peserta = ".$nomor_peserta;
echo "<br>";
echo "paud = ".$paud;
echo "<br>";
echo "tk = ".$tk;
echo "<br>";
echo "skhun = ".$skhun;
echo "<br>";
echo "ijazah = ".$ijazah;
echo "<br>";
echo "hobi = ".$hobi;
echo "<br>";
echo "cita_cita = ".$cita_cita;
echo "<br>";
echo "Nama = ".$nama;
echo "<br>";
echo " jenis_kelamin = ".$jenis_kelamin;
echo "<br>";
echo " nisn = ".$nisn;
echo "<br>";
echo " nik = ".$nik;
echo "<br>";
echo " tempat_lahir = ".$tempat_lahir;
echo "<br>";
echo " tanggal_lahir = ".$tanggal_lahir;
echo "<br>";
echo " agama = ".$agama;
echo "<br>";
echo " berkebutuhan_khusus = ".$berkebutuhan_khusus;
echo "<br>";
echo " alamat = ".$alamat;
echo "<br>";
echo " rt = ".$rt;
echo "<br>";
echo " rw = ".$rw;
echo "<br>";
echo " dusun = ".$dusun;
echo "<br>";
echo " kelurahan = ".$kelurahan;
echo "<br>";
echo " kecamatan = ".$kecamatan;
echo "<br>";
echo " kode_pos = ".$kode_pos;
echo "<br>";
echo " tempat_tinggal = ".$tempat_tinggal;
echo "<br>";
echo " transportasi = ".$transportasi;
echo "<br>";
echo " no_hp = ".$no_hp;
echo "<br>";
echo " no_telp = ".$no_telp;
echo "<br>";
echo "email = ".$email;
echo "<br>";
echo "penerima_kps = ".$penerima_kps;
echo "<br>";
echo "no_kps = ".$no_kps;
echo "<br>";
echo "no_kps = ".$kewarganegaraan;

include "koneksi1.php";
//query SQL utk insert into
$query = "INSERT INTO pendaftaran SET jenis_pendaftaran='$jenis_pendaftaran', tanggal_masuk='$tanggal_masuk', nis='$nis', nomor_peserta='$nomor_peserta', paud='$paud', tk='$tk', no_skhun='$skhun', no_ijazah='$ijazah', hobi='$hobi', cita_cita='$cita_cita', nama='$nama', jenis_kelamin='$jenis_kelamin', nisn='$nisn', nik='$nik', tempat_lahir='$tempat_lahir', tanggal_lahir='$tanggal_lahir', agama='$agama', berkebutuhan_khusus='$berkebutuhan_khusus', alamat='$alamat', rt='$rt', rw='$rw', dusun='$dusun', kelurahan='$kelurahan', kecamatan='$kecamatan', kode_pos='$kode_pos', tempat_tinggal='$tempat_tinggal', transportasi='$transportasi', no_hp='$no_hp', no_telp='$no_telp', email='$email', penerima_kps='$penerima_kps', no_kps='$no_kps', kewarganegaraan='$kewarganegaraan'";
mysqli_query($conn, $query);
 ?>
