<!DOCTYPE html>
<html lang="en" dir="ltr">
  <head>
    <!-- Referensi bootstrap  -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="
    sha284-MCw98/SFnGE8fJT3GXwE0ngsV72t27NXFoaoApmYm8liuXoPkF0JwJ8RdknLPMO" crossorigin="anonymus">
    <style>
      .warning{color: #FF0000;}
    </style>
  </head>
  <body>
      <?php
      //pernyataan variabel error
      $error_jenis_pendaftaran="";
      $error_tanggal_masuk="";
      $error_nis="";
      $error_nomor_peserta="";
      $error_paud="";
      $error_tk="";
      $error_skhun="";
      $error_ijazah="";
      $error_hobi="";
      $error_cita_cita="";
      $error_nama="";
      $error_jenis_kelamin="";
      $error_nisn="";
      $error_nik="";
      $error_tempat_lahir="";
      $error_tanggal_lahir="";
      $error_agama="";
      $error_berkebutuhan_khusus="";
      $error_alamat="";
      $error_rt="";
      $error_rw="";
      $error_dusun="";
      $error_kelurahan="";
      $error_kecamatan="";
      $error_kode_pos="";
      $error_tempat_tinggal="";
      $error_transportasi="";
      $error_no_hp="";
      $error_no_telp="";
      $error_email="";
      $error_penerima_kps="";
      $error_no_kps="";
      $error_kewarganegaraan="";

      //pernyataan variabel
      $jenis_pendaftaran="";
      $tanggal_masuk="";
      $nis="";
      $nomor_peserta="";
      $paud="";
      $tk="";
      $skhun="";
      $ijazah="";
      $hobi="";
      $cita_cita="";
      $nama="";
      $jenis_kelamin="";
      $nisn="";
      $nik="";
      $tempat_lahir="";
      $tanggal_lahir="";
      $agama="";
      $berkebutuhan_khusus="";
      $alamat="";
      $rt="";
      $rw="";
      $dusun="";
      $kelurahan="";
      $kecamatan="";
      $kode_pos="";
      $tempat_tinggal="";
      $transportasi="";
      $no_hp="";
      $no_telp="";
      $email="";
      $penerima_kps="";
      $no_kps="";
      $kewarganegaraan="";

      //pengecekan isian data dan penampilan pesan error
      if ($_SERVER["REQUEST_METHOD"] == "POST") {
      if (empty($_POST["jenis_pendaftaran"]))
      {
        $error_jenis_pendaftaran= "Jenis Pendaftaran tidak boleh kosong";
      }
      else
      {
        $jenis_pendaftaran= cek_input($_POST["jenis_pendaftaran"]);
      }

        if (empty($_POST["tanggal_masuk"]))
        {
          $error_tanggal_masuk= "Tanggal masuk tidak boleh kosong";
        }
        else
        {
          $tanggal_masuk= cek_input($_POST["tanggal_masuk"]);
          if (!is_numeric($tanggal_masuk))
          {
            $error_tanggal_masuk="Tanggal masuk hanya boleh angka";
          }
        }

        if (empty($_POST["nis"]))
        {
          $error_nis= "NIS tidak boleh kosong";
        }
        else
        {
          $nis= cek_input($_POST["nis"]);
          if (!is_numeric($nis))
          {
            $error_nis="NIS hanya boleh angka";
          }
        }

        if (empty($_POST["nomor_peserta"]))
        {
          $error_nomor_peserta= "Nomor Peserta tidak boleh kosong";
        }
        else
        {
          $nomor_peserta= cek_input($_POST["nomor_peserta"]);
          if (!is_numeric($nomor_peserta))
          {
            $error_nomor_peserta="Nomor Peserta hanya boleh angka";
          }
        }

        if (empty($_POST["paud"]))
        {
          $error_paud= "Pilihan tidak boleh kosong";
        }
        else
        {
          $paud= cek_input($_POST["paud"]);
        }

        if (empty($_POST["tk"]))
        {
          $error_tk= "Pilihan tidak boleh kosong";
        }
        else
        {
          $tk= cek_input($_POST["tk"]);
        }

        if (empty($_POST["skhun"]))
        {
          $error_skhun= "No SKHUN tidak boleh kosong";
        }
        else
        {
          $skhun= cek_input($_POST["skhun"]);
          if (!is_numeric($skhun))
          {
            $error_skhun="No SKHUN hanya boleh angka";
          }
        }

        if (empty($_POST["ijazah"]))
        {
          $error_ijazah= "No ijazah tidak boleh kosong";
        }
        else
        {
          $ijazah= cek_input($_POST["ijazah"]);
          if (!is_numeric($ijazah))
          {
            $error_ijazah="No ijazah hanya boleh angka";
          }
        }

        if (empty($_POST["hobi"]))
        {
          $error_hobi= "Jenis Pendaftaran tidak boleh kosong";
        }
        else
        {
          $hobi= cek_input($_POST["hobi"]);
        }

        if (empty($_POST["cita_cita"]))
        {
          $error_cita_cita= "Jenis Pendaftaran tidak boleh kosong";
        }
        else
        {
          $cita_cita= cek_input($_POST["cita_cita"]);
        }


          if (empty($_POST["nama"]))
          {
            $error_nama= "Nama tidak boleh kosong";
          }
          else
          {
            $nama= cek_input($_POST["nama"]);
            if (!preg_match("/^[a-zA-Z]*$/",$nama))
            {
              $nameErr="Inputan hanya boleh huruf dan spasi";
            }
          }

          if (empty($_POST["jenis_kelamin"]))
          {
            $error_jenis_kelamin= "Pilihan tidak boleh kosong";
          }
          else
          {
            $jenis_kelamin= cek_input($_POST["jenis_kelamin"]);
          }

          if (empty($_POST["nisn"]))
          {
            $error_nisn= "No NISN tidak boleh kosong";
          }
          else
          {
            $nisn= cek_input($_POST["nisn"]);
            if (!is_numeric($nisn))
            {
              $error_nisn="No NISN hanya boleh angka";
            }
          }

          if (empty($_POST["nik"]))
          {
            $error_nik= "No NIK tidak boleh kosong";
          }
          else
          {
            $nik= cek_input($_POST["nik"]);
            if (!is_numeric($nik))
            {
              $error_nik="No NIK hanya boleh angka";
            }
          }

          if (empty($_POST["tempat_lahir"]))
          {
            $error_tempat_lahir= "tempat_lahir tidak boleh kosong";
          }
          else
          {
            $tempat_lahir = cek_input($_POST["tempat_lahir"]);
          }

          if (empty($_POST["tanggal_lahir"]))
          {
            $error_tanggal_lahir= "Tanggal lahir tidak boleh kosong";
          }
          else
          {
            $tanggal_lahir= cek_input($_POST["tanggal_lahir"]);
            if (!is_numeric($tanggal_lahir))
            {
              $error_tanggal_lahir="Tanggal lahir hanya boleh angka";
            }
          }

          if (empty($_POST["agama"]))
          {
            $error_agama= "Agamaa tidak boleh kosong";
          }
          else
          {
            $agama= cek_input($_POST["agama"]);
          }

          if (empty($_POST["berkebutuhan_khusus"]))
          {
            $error_berkebutuhan_khusus= "Pilihan tidak boleh kosong";
          }
          else
          {
            $berkebutuhan_khusus= cek_input($_POST["berkebutuhan_khusus"]);
          }

          if (empty($_POST["alamat"]))
          {
            $error_alamat= "Alamat tidak boleh kosong";
          }
          else
          {
            $alamat = cek_input($_POST["alamat"]);
          }

          if (empty($_POST["rt"]))
          {
            $error_rt= "No RT tidak boleh kosong";
          }
          else
          {
            $rt= cek_input($_POST["rt"]);
            if (!is_numeric($rt))
            {
              $error_rt="No RT hanya boleh angka";
            }
          }

          if (empty($_POST["rw"]))
          {
            $error_rw= "No RW tidak boleh kosong";
          }
          else
          {
            $rw= cek_input($_POST["rw"]);
            if (!is_numeric($rw))
            {
              $error_rw="No RW hanya boleh angka";
            }
          }

          if (empty($_POST["dusun"]))
          {
            $error_dusun= "Dusun tidak boleh kosong";
          }
          else
          {
            $dusun = cek_input($_POST["dusun"]);
          }

          if (empty($_POST["kelurahan"]))
          {
            $error_kelurahan= "Kelurahan tidak boleh kosong";
          }
          else
          {
            $kelurahan = cek_input($_POST["kelurahan"]);
          }

          if (empty($_POST["kecamatan"]))
          {
            $error_kecamatan= "Kecamatan tidak boleh kosong";
          }
          else
          {
            $kecamatan = cek_input($_POST["kecamatan"]);
          }

          if (empty($_POST["kode_pos"]))
          {
            $error_kode_pos= "No kode pos tidak boleh kosong";
          }
          else
          {
            $kode_pos= cek_input($_POST["kode_pos"]);
            if (!is_numeric($kode_pos))
            {
              $error_kode_pos="No kode pos hanya boleh angka";
            }
          }

          if (empty($_POST["tempat_tinggal"]))
          {
            $error_tempat_tinggal= "Tempat Tinggal tidak boleh kosong";
          }
          else
          {
            $tempat_tinggal= cek_input($_POST["tempat_tinggal"]);
          }

          if (empty($_POST["transportasi"]))
          {
            $error_transportasi= "Modal Transportasi tidak boleh kosong";
          }
          else
          {
            $transportasi= cek_input($_POST["transportasi"]);
          }

          if (empty($_POST["no_hp"]))
          {
            $error_no_hp= "Nomor HP tidak boleh kosong";
          }
          else
          {
            $no_hp= cek_input($_POST["no_hp"]);
            if (!is_numeric($no_hp))
            {
              $error_no_hp="Nomor HP hanya boleh angka";
            }
          }

          if (empty($_POST["no_telp"]))
          {
            $error_no_telp= "Nomor Telepon tidak boleh kosong";
          }
          else
          {
            $no_telp= cek_input($_POST["no_telp"]);
            if (!is_numeric($no_telp))
            {
              $error_no_telp="Nomor Telepon hanya boleh angka";
            }
          }

        if (empty($_POST["email"]))
        {
          $error_email= "Email tidak boleh kosong";
        }
        else
        {
          $email= cek_input($_POST["email"]);
          if (!filter_var($email, FILTER_VALIDATE_EMAIL))
          {
            $error_email="Format email invalid";
          }
        }

        if (empty($_POST["penerima_kps"]))
        {
          $error_penerima_kps= "Pilihan tidak boleh kosong";
        }
        else
        {
          $penerima_kps= cek_input($_POST["penerima_kps"]);
        }

        if (empty($_POST["no_kps"]))
        {
          $error_no_kps= "No KPS/KPH/KIP tidak boleh kosong";
        }
        else
        {
          $no_kps= cek_input($_POST["no_kps"]);
          if (!is_numeric($no_kps))
          {
            $error_no_kps="No KPS/KPH/KIP hanya boleh angka";
          }
        }

        if (empty($_POST["kewarganegaraan"]))
        {
          $error_kewarganegaraan= "Pilihan tidak boleh kosong";
        }
        else
        {
          $kewarganegaraan= cek_input($_POST["kewarganegaraan"]);
        }
      }

      // function pengecekan data
      function cek_input($data){
        $data= trim($data);
        $data= stripslashes($data);
        $data= htmlspecialchars($data);
        return $data;
      }
       ?>

       <div class="row">
         <div class="col-md-6">
           <div class="card">
             <div class="card-header">
               Validasi Form dengan PHP
             </div>
             <div class="card-body">
               <!-- Pembuatan form  -->
               <form action="simpan.php" method="post">

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="jenis_pendaftaran" class="col-sm-2 col-form-tabel">Jenis Pendaftaran</label>
                   <div class="col-sm-10">
                     <select name="jenis_pendaftaran" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Siswa Baru">Siswa Baru</option>
                       <option value="Pindahan">Pindahan</option>
                     </select><span class="warning"><?php echo $error_jenis_pendaftaran; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="tanggal_masuk" class="col-sm-2 col-form-label">Tanggal Masuk</label>
                   <div class="col-sm-10">
                     <input type="text" name="tanggal_masuk" class="form-control <?php echo ($error_tanggal_masuk !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_masuk" placeholder="01012021" value="<?php echo $tanggal_masuk; ?>"><span class="warning"><?php echo $error_tanggal_masuk; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="nis" class="col-sm-2 col-form-label">NIS</label>
                   <div class="col-sm-10">
                     <input type="text" name="nis" class="form-control <?php echo ($error_nis !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_masuk" placeholder="01012021" value="<?php echo $nis; ?>"><span class="warning"><?php echo $error_nis; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="nomor_peserta" class="col-sm-2 col-form-label">Nomor Peserta Ujian</label>
                   <div class="col-sm-10">
                     <input type="text" name="nomor_peserta" class="form-control <?php echo ($error_nomor_peserta !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_masuk" placeholder="01012021" value="<?php echo $nomor_peserta; ?>"><span class="warning"><?php echo $error_nomor_peserta; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="paud" class="col-sm-2 col-form-label">Apakah pernah PAUD?</label>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="paud" value="Ya">Ya
                   </div>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="paud" value="Tidak">Tidak
                   </div>
                   <div class="col-sm-10">
                     <span class="warning"><?php echo $error_paud; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="tk" class="col-sm-2 col-form-label">Apakah pernah TK?</label>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="tk" value="Ya">Ya
                   </div>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="tk" value="Tidak">Tidak
                   </div>
                   <div class="col-sm-10">
                     <span class="warning"><?php echo $error_tk; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="skhun" class="col-sm-2 col-form-label">No SKHUN Sebelumnya</label>
                   <div class="col-sm-10">
                     <input type="text" name="skhun" class="form-control <?php echo ($error_skhun !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_masuk" placeholder="01012021" value="<?php echo $skhun; ?>"><span class="warning"><?php echo $error_skhun; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="ijazah" class="col-sm-2 col-form-label">No ijazah Sebelumnya</label>
                   <div class="col-sm-10">
                     <input type="text" name="ijazah" class="form-control <?php echo ($error_ijazah !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_masuk" placeholder="01012021" value="<?php echo $ijazah; ?>"><span class="warning"><?php echo $error_ijazah; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="hobi" class="col-sm-2 col-form-tabel">Hobi</label>
                   <div class="col-sm-10">
                     <select name="hobi" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Olah Raga">Olah Raga</option>
                       <option value="Kesenian">Kesenian</option>
                       <option value="Membaca">Membaca</option>
                       <option value="Menulis">Menulis</option>
                       <option value="Traveling">Traveling</option>
                       <option value="Lainnya">Lainnya</option>
                     </select><span class="warning"><?php echo $error_hobi; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="cita_cita" class="col-sm-2 col-form-tabel">Cita-Cita</label>
                   <div class="col-sm-10">
                     <select name="cita_cita" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="PNS">PNS</option>
                       <option value="TNI/POLRI">TNI/POLRI</option>
                       <option value="Guru/Dosen">Guru/Dosen</option>
                       <option value="Dokter">Dokter</option>
                       <option value="Politikus">Politikus</option>
                       <option value="Wiraswasta">Wiraswasta</option>
                       <option value="Seni/Lukis">Seni/Lukis</option>
                       <option value="Lainnya">Lainnya</option>
                     </select><span class="warning"><?php echo $error_cita_cita; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="nama" class="col-sm-2 col-form-tabel">Nama Lengkap</label>
                   <div class="col-sm-10">
                     <input type="text" name="nama" class="form-control" <?php echo ($error_nama !="" ? "is-invalid" : ""); ?>
                     id="nama" placeholder="Nama" value="<?php echo $nama; ?>"><span class="warning"><?php echo $error_nama; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="jenis_kelamin" class="col-sm-2 col-form-label">Jenis Kelamin</label>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="jenis_kelamin" value="L">Laki
                   </div>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="jenis_kelamin" value="P">Perempuan
                   </div>
                   <div class="col-sm-10">
                     <span class="warning"><?php echo $error_jenis_kelamin; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="nisn" class="col-sm-2 col-form-label">No NISN</label>
                   <div class="col-sm-10">
                     <input type="text" name="nisn" class="form-control <?php echo ($error_nisn !="" ? "is-invalid" : ""); ?>"
                     id="nisn" placeholder="01012021" value="<?php echo $nisn; ?>"><span class="warning"><?php echo $error_nisn; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="nik" class="col-sm-2 col-form-label">No NIK</label>
                   <div class="col-sm-10">
                     <input type="text" name="nik" class="form-control <?php echo ($error_nik !="" ? "is-invalid" : ""); ?>"
                     id="nik" placeholder="01012021" value="<?php echo $nisn; ?>"><span class="warning"><?php echo $error_nik; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="tempat_lahir" class="col-sm-2 col-form-tabel">Tempat Lahir</label>
                   <div class="col-sm-10">
                     <input type="text" name="tempat_lahir" class="form-control" <?php echo ($error_tempat_lahir !="" ? "is-invalid" : ""); ?>
                     id="tempat_lahir" placeholder="Tempat Lahir" value="<?php echo $tempat_lahir; ?>"><span class="warning"><?php echo $error_tempat_lahir; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="tanggal_lahir" class="col-sm-2 col-form-label">Tanggal Lahir </label>
                   <div class="col-sm-10">
                     <input type="text" name="tanggal_lahir" class="form-control <?php echo ($error_tanggal_lahir !="" ? "is-invalid" : ""); ?>"
                     id="tanggal_lahir" placeholder="01012021" value="<?php echo $tanggal_lahir; ?>"><span class="warning"><?php echo $error_tanggal_lahir; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="agama" class="col-sm-2 col-form-tabel">Agama</label>
                   <div class="col-sm-10">
                     <select name="agama" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Islam">Islam</option>
                       <option value="Kristen/Protestan">Kristen/Protestan</option>
                       <option value="Katholik">Katholik</option>
                       <option value="Hindu">Hindu</option>
                       <option value="Budha">Budha</option>
                       <option value="Kong Hu Chu">Kong Hu Chu</option>
                       <option value="Lainnya">Lainnya</option>
                     </select><span class="warning"><?php echo $error_agama; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="berkebutuhan_khusus" class="col-sm-2 col-form-tabel">Berkebutuhan Khusus</label>
                   <div class="col-sm-10">
                     <select name="berkebutuhan_khusus" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Tidak">Tidak</option>
                       <option value="Netra">Netra</option>
                       <option value="Rungu">Rungu</option>
                       <option value="Grahita Ringan">Grahita Ringan</option>
                       <option value="Grahita Sedang">Grahita Sedang</option>
                       <option value="Daksa Ringan">Daksa Ringan</option>
                       <option value="Daksa Sedang">Daksa Sedang</option>
                       <option value="Laras">Laras</option>
                       <option value="Wicara">Wicara</option>
                       <option value="Tuna ganda">Tuna ganda</option>
                       <option value="Hiper aktif">Hiper aktif</option>
                       <option value="Cerdas Istimewa">Cerdas Istimewa</option>
                       <option value="Bakat Istimewa">Bakat Istimewa</option>
                       <option value="Kesulitan belaja">Kesulitan belajar</option>
                       <option value="Narkoba">Narkoba</option>
                       <option value="Indigo">Indigo</option>
                       <option value="Down Sindrome">Down Sindrome</option>
                       <option value="Autis">Autis</option>
                     </select><span class="warning"><?php echo $error_berkebutuhan_khusus; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="alamat" class="col-sm-2 col-form-tabel">Alamat</label>
                   <div class="col-sm-10">
                     <input type="text" name="alamat" class="form-control" <?php echo ($error_alamat !="" ? "is-invalid" : ""); ?>
                     id="alamat" placeholder="Alamat" value="<?php echo $alamat; ?>"><span class="warning"><?php echo $error_alamat; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="rt" class="col-sm-2 col-form-label">RT</label>
                   <div class="col-sm-10">
                     <input type="text" name="rt" class="form-control <?php echo ($error_rt !="" ? "is-invalid" : ""); ?>"
                     id="rt" placeholder="01012021" value="<?php echo $rt; ?>"><span class="warning"><?php echo $error_rt; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="rw" class="col-sm-2 col-form-label">RW</label>
                   <div class="col-sm-10">
                     <input type="text" name="rw" class="form-control <?php echo ($error_rw !="" ? "is-invalid" : ""); ?>"
                     id="rw" placeholder="01012021" value="<?php echo $rw; ?>"><span class="warning"><?php echo $error_rw; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="dusun" class="col-sm-2 col-form-tabel">Nama Dusun</label>
                   <div class="col-sm-10">
                     <input type="text" name="dusun" class="form-control" <?php echo ($error_dusun !="" ? "is-invalid" : ""); ?>
                     id="dusun" placeholder="Dusun" value="<?php echo $dusun; ?>"><span class="warning"><?php echo $error_dusun; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="kelurahan" class="col-sm-2 col-form-tabel">Nama Kelurahan/Desa</label>
                   <div class="col-sm-10">
                     <input type="text" name="kelurahan" class="form-control" <?php echo ($error_kelurahan !="" ? "is-invalid" : ""); ?>
                     id="kelurahan" placeholder="Kelurahan" value="<?php echo $kelurahan; ?>"><span class="warning"><?php echo $error_kelurahan; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="kecamatan" class="col-sm-2 col-form-tabel">Nama Kecamatan</label>
                   <div class="col-sm-10">
                     <input type="text" name="kecamatan" class="form-control" <?php echo ($error_kecamatan !="" ? "is-invalid" : ""); ?>
                     id="kecamatan" placeholder="Kecamatan" value="<?php echo $kecamatan; ?>"><span class="warning"><?php echo $error_kecamatan; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="kode_pos" class="col-sm-2 col-form-label">Kode Pos</label>
                   <div class="col-sm-10">
                     <input type="text" name="kode_pos" class="form-control <?php echo ($error_kode_pos !="" ? "is-invalid" : ""); ?>"
                     id="kode_pos" placeholder="01012021" value="<?php echo $kode_pos; ?>"><span class="warning"><?php echo $error_kode_pos; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="tempat_tinggal" class="col-sm-2 col-form-tabel">Tempat Tinggal</label>
                   <div class="col-sm-10">
                     <select name="tempat_tinggal" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Bersama Orang Tua">Bersama Orang Tua</option>
                       <option value="Wali">Wali</option>
                       <option value="Kos">Kos</option>
                       <option value="Asrama">Asrama</option>
                       <option value="Lainnya">Lainnya</option>
                     </select><span class="warning"><?php echo $error_tempat_tinggal; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="transportasi" class="col-sm-2 col-form-tabel">Modal Transportasi</label>
                   <div class="col-sm-10">
                     <select name="transportasi" class="form-control">
                       <option selected disabled value="">Pilih...</option>
                       <option value="Jalan Kaki">Jalan Kaki</option>
                       <option value="Kendaraan Pribadi">Kendaraan Pribadi</option>
                       <option value="Kendaraan Umum">Kendaraan Umum</option>
                       <option value="Jemputan Sekolah">Jemputan Sekolah</option>
                       <option value="Kereta Api">Kereta Api</option>
                       <option value="Ojek">Ojek</option>
                       <option value="Becak">Becak</option>
                       <option value="Rakit">Rakit</option>
                       <option value="Lainnya">Lainnya</option>
                     </select><span class="warning"><?php echo $error_transportasi; ?></sjpan>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="no_hp" class="col-sm-2 col-form-label">No HP</label>
                   <div class="col-sm-10">
                     <input type="text" name="no_hp" class="form-control <?php echo ($error_no_hp !="" ? "is-invalid" : ""); ?>"
                     id="no_hp" placeholder="08123456789" value="<?php echo $no_hp; ?>"><span class="warning"><?php echo $error_no_hp; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="no_telp" class="col-sm-2 col-form-label">No telp</label>
                   <div class="col-sm-10">
                     <input type="text" name="no_telp" class="form-control <?php echo ($error_no_telp !="" ? "is-invalid" : ""); ?>"
                     id="no_telp" placeholder="031123546789" value="<?php echo $no_telp; ?>"><span class="warning"><?php echo $error_no_telp; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="email" class="col-sm-2 col-form-label">Email Pribadi</label>
                   <div class="col-sm-10">
                     <input type="text" name="email" class="form-control <?php echo ($error_email !="" ? "is-invalid" : ""); ?>"
                     id="email" placeholder="email" value="<?php echo $email; ?>"><span class="warning"><?php echo $error_email; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="penerima_kps" class="col-sm-2 col-form-label">Penerima KPS/PKH/KIP</label>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="penerima_kps" value="Ya">Ya
                   </div>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="penerima_kps" value="Tidak">Tidak
                   </div>
                   <div class="col-sm-10">
                     <span class="warning"><?php echo $error_penerima_kps; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="no_kps" class="col-sm-2 col-form-label">No KPS/PKH/KIP</label>
                   <div class="col-sm-10">
                     <input type="text" name="no_kps" class="form-control <?php echo ($error_no_kps !="" ? "is-invalid" : ""); ?>"
                     id="no_kps" placeholder="01012021" value="<?php echo $no_kps; ?>"><span class="warning"><?php echo $error_no_kps; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan input data web  -->
                 <div class="form-group row">
                   <label for="kewarganegaraan" class="col-sm-2 col-form-label">Kewarganegaraan</label>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="kewarganegaraan" value="WNI">WNI
                   </div>
                   <div class="form-check form-check-inline">
                     <input class="form-check-input" type="radio" name="kewarganegaraan" value="WNA">WNA
                   </div>
                   <div class="col-sm-10">
                     <span class="warning"><?php echo $error_kewarganegaraan; ?></span>
                   </div>
                 </div>

                 <!-- Pembuatan button submit dan mengimpor data ke excel -->
                 <div class="form-group row">
                   <div class="col-sm-10">
                     <button type="submit" class="btn btn-primary">Submit</button>
                   </div>
                 </div>
               </form>
               <a href="reportpendaftaran.php"><button type="submit" class="btn btn-primary">Tampilkan Data di Excel</button></a>
             </div>
           </div>
         </div>
       </div>
  </body>
</html>
