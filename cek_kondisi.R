cek_persen_tes = function(persen_tes, n_data_tes = NULL){
  persen_tes_desimal = persen_tes / 100
  stopifnot(
    "Persentase data ujian harus lebih dari nol dan minimal menghasilkan
    satu buah titik data untuk diuji, tapi tidak boleh melebihi 100." =
      persen_tes_desimal > 0 & persen_tes_desimal <= 1
  )

  stopifnot("Persentase data ujian yang Anda pilih terlalu kecil, sehingga tidak
            menghasilkan titik data untuk ujian bagi model. Coba tingkatkan
            persentase data ujian Anda sehingga  menghasilkan minimal satu buah
            titik data untuk ujian." = n_data_tes >= 1)

}
