fpaket = function(paket1, ...) {
  nama_paket = c(paket1, ...)
  for (k in seq_along(nama_paket)) {
    if (isFALSE(suppressWarnings(require(nama_paket[k],
      character.only = TRUE
    )))) {
      install.packages(nama_paket[k], dependencies = TRUE,
                       repos = "https://cran.usk.ac.id/")
      library(nama_paket[k], character.only = TRUE)
    }
  }
}

fpaket("openxlsx", "data.table", "minpack.lm", "ggplot2", "investr")

# Contoh nama_file = "contoh data penduduk banyak kota.xlsx"

fbaca = function(nama_file, sheet_terpilih = NULL) {
  semua_sheet = if (is.null(sheet_terpilih)) getSheetNames(nama_file) else sheet_terpilih
  list_luaran_fbaca = lapply(semua_sheet, function(k) {
    tabel = read.xlsx(nama_file, sheet = k) |> as.data.table()
    for (j in 1:2) set(tabel, j = j, value = as.integer(tabel[[j]]))
    return(tabel)
  })
  names(list_luaran_fbaca) = semua_sheet
  return(list_luaran_fbaca)
}


.flatih = function(tabel_data,
                    jenis_model = c(
                      "semua",
                      "aritmatik", "eksponensial", "geometrik", "logistik"
                    )) {
  tabel_data$urutan = tabel_data$tahun - tabel_data$tahun[1]
  tabel_data = tabel_data[, c("tahun", "urutan", "penduduk")]
  penduduk_awal = tabel_data$penduduk[1]
  penduduk_akhir = tabel_data$penduduk[nrow(tabel_data)]
  tebakan_laju = log(penduduk_akhir / penduduk_awal) * 1 / nrow(tabel_data)
  mod_aritmatik = eval(bquote(nls(penduduk ~ m_lin * urutan + c_lin,
    data = .(tabel_data),
    start = list(
      m_lin = (penduduk_akhir - penduduk_awal) / nrow(tabel_data),
      c_lin = penduduk_awal
    )
  )))
  mod_eksponensial = eval(bquote(nls(penduduk ~ a_eks * exp(r_eks * urutan),
    data = .(tabel_data),
    start = list(
      a_eks = penduduk_awal,
      r_eks = tebakan_laju
    )
  )))
  mod_geometrik = eval(bquote(nls(penduduk ~ a_geo * (1 + r_geo)^urutan,
    data = .(tabel_data),
    start = list(
      a_geo = tabel_data$penduduk[1],
      r_geo = tebakan_laju
    )
  )))
  mod_logistik = eval(bquote(nlsLM(penduduk ~ k_log / (1 + a_log * exp(-r_log * urutan)),
    data = .(tabel_data),
    start = list(
      k_log = 1.5 * max(tabel_data$penduduk),
      a_log = (1.5 * max(tabel_data$penduduk) /
        tabel_data$penduduk[1]) - 1,
      r_log = tebakan_laju
    ), lower = c(
      k_log = max(tabel_data$penduduk),
      a_log = (1.1 * max(tabel_data$penduduk) /
        penduduk_awal) - 1,
      r_log = 0.1 * tebakan_laju
    ),
    upper = c(
      k_log = 100 * max(tabel_data$penduduk),
      a_log = (100 * max(tabel_data$penduduk) /
        penduduk_awal) - 1,
      r_log = 0.5 * tebakan_laju
    ),
    control = nls.lm.control(maxiter = 200, maxfev = 500)
  )))
  jenis = match.arg(jenis_model)
  list_hasil_flatih = switch(jenis,
    "aritmatik" = mod_aritmatik,
    "eksponensial" = mod_eksponensial,
    "geometrik" = mod_geometrik,
    "logistik" = mod_logistik,
    "semua" = list(
      aritmatik = mod_aritmatik,
      eksponensial = mod_eksponensial,
      geometrik = mod_geometrik,
      logistik = mod_logistik
    )
  )

  return(list_hasil_flatih)
}


.ftes = function(tabel_data_latih, tabel_data_uji) {
  data_asal = rbind(tabel_data_latih, tabel_data_uji)
  hasil_model = .flatih(tabel_data_latih)
  parameter = lapply(
    hasil_model,
    function(n) {
      par = as.data.frame(coef(n))
      data.table(
        nama_param = rownames(par),
        nilai_param = round(par[[1]], 4)
      )
    }
  )
  tabel_parameter = do.call(rbind.data.frame, parameter)
  tabel_parameter_model = cbind(
    model = rownames(tabel_parameter),
    tabel_parameter
  )

  prediksi = lapply(
    hasil_model,
    function(model) predict(model, newdata = tabel_data_uji)) |>
    do.call(cbind.data.frame, args = _) |>
    setNames(paste0("prediksi_", names(hasil_model)))

  selisih = lapply(
    prediksi,
    function(pred) (pred - tabel_data_uji$penduduk) |> round(2)
  ) |>
    do.call(cbind.data.frame, args = _) |>
    setnames(paste0("selisih_", names(hasil_model)))
  gabungan = cbind(tabel_data_uji, prediksi, selisih)
  akurasi = data.table(
    model = names(hasil_model),
    mape = apply(
      selisih, 2,
      function(kolom) 100 * mean(abs(kolom) / tabel_data_uji$penduduk)
    ),
    mae = apply(selisih, 2, function(kolom) mean(abs(kolom))),
    rss = apply(selisih, 2, function(kolom) sum(kolom^2))
   ,
    row.names = NULL, check.names = FALSE
  )
  terkecil = data.table(
    mape_terkecil = akurasi$model[which.min(akurasi$mape)],
    mae_terkecil = akurasi$model[which.min(akurasi$mae)],
    rss_terkecil = akurasi$model[which.min(akurasi$rss)]
  )
  terbaik = terkecil[1, ] |>
    as.character() |>
    table() |>
    which.max() |>
    names() |>
    as.data.table() |>
    setnames("model_terbaik")

  model_terbaik_data_lengkap = .flatih(data_asal, terbaik$model_terbaik)

  list_hasil_ftes = list(
    data_semua = data_asal,
    tabel_parameter = tabel_parameter_model,
    prediksi_tes = gabungan,
    akurasi_model = akurasi,
    model_dengan_galat_terkecil = terkecil,
    jenis_model_terbaik = terbaik,
    parameter_model_terbaik_semua_data = model_terbaik_data_lengkap
  )
  return(list_hasil_ftes)
}

# Input fmodel adalah output dari fbaca(nama_file, sheet terpilih)

fmodel = function(list_hasil_fbaca, sheet_terpilih = NULL, persen_tes = 20) {
  cek_persen_tes(persen_tes)
  semua_sheet = if (is.null(sheet_terpilih)) names(list_hasil_fbaca) else sheet_terpilih
  data_semua = data_latihan = data_tes = hasil_uji_model =
    vector("list", length(semua_sheet))
  names(data_semua) = semua_sheet
  nbaris = ntes = vector("numeric", length(semua_sheet))
  for (k in seq_along(semua_sheet)) {
    data_semua[[k]] = list_hasil_fbaca[[k]]
    data_semua[[k]] =  data_semua[[k]] |>
      transform(urutan = tahun - tahun[1])
   nbaris[k] = nrow(data_semua[[k]])
    ntes[k] = round(0.01 * persen_tes * nbaris[k])
    cek_persen_tes(persen_tes, ntes[k])
    data_latihan[[k]] = if ((nbaris[k] - ntes[k]) < 1) {
      data_semua[[k]]
    } else {
      head(data_semua[[k]], nbaris[k] - ntes[k])
    }
    data_tes[[k]] = tail(data_semua[[k]], ntes[k])
    names(data_latihan) = names(data_tes) = semua_sheet
    hasil_uji_model[[k]] = .ftes(data_latihan[[k]], data_tes[[k]])
  }
  names(hasil_uji_model) = names(data_semua) = semua_sheet
  hasil_fmodel_semua = hasil_uji_model
  return(hasil_fmodel_semua)
}

# Input fproyeksi() adalah output fmodel()

fproyeksi = function(list_hasil_fmodel, tahun_proyeksi,
                     sheet_terpilih = NULL,
                      persen_tes = 20) {
  nama_sheet = if (is.null(sheet_terpilih)) names(list_hasil_fmodel) else sheet_terpilih
  n_sheet = length(nama_sheet)
  hasil_pilih_model = list_hasil_fmodel
  n_tahun_proyeksi = length(tahun_proyeksi)
  tahun_awal_model = numeric(n_sheet)
  urutan_tahun_proyeksi = vector("list", n_sheet)
  tabel_hasil_fproyeksi = vector("list", n_sheet)
  prediksi_interval = vector("list", n_sheet)
  modterbaik = vector("list", n_sheet)

  for (k in seq_len(n_sheet)) {
    tahun_awal_model[k] = hasil_pilih_model[[k]]$data_semua$tahun[1]
    urutan_tahun_proyeksi[[k]] = tahun_proyeksi - tahun_awal_model[k]
    modterbaik[[k]] = hasil_pilih_model[[k]]$parameter_model_terbaik_semua_data
    names(modterbaik) = nama_sheet
    tabel_hasil_fproyeksi[[k]] = data.table(
      tahun = tahun_proyeksi,
      urutan = urutan_tahun_proyeksi[[k]],
      penduduk = rep(NA_real_, n_tahun_proyeksi),
      batas_bawah = rep(NA_real_, n_tahun_proyeksi),
      batas_atas = rep(NA_real_, n_tahun_proyeksi)
    )
    tabel_hasil_fproyeksi[[k]]$penduduk = round(predict(
      modterbaik[[k]],
      newdata = data.table(urutan = urutan_tahun_proyeksi[[k]])
    ))
    try({
      prediksi_interval[[k]] = predFit(
        modterbaik[[k]],
        newdata = data.table(urutan = urutan_tahun_proyeksi[[k]]),
        interval = "prediction"
      )
      tabel_hasil_fproyeksi[[k]]$batas_bawah = round(prediksi_interval[[k]][, 2])
      tabel_hasil_fproyeksi[[k]]$batas_atas = round(prediksi_interval[[k]][, 3])
    })
    names(tabel_hasil_fproyeksi)[k] = paste(
      "Proyeksi Model Terbaik untuk",
      nama_sheet[k]
    )
    hasil_pilih_model[[k]]$proyeksi_model_terbaik = tabel_hasil_fproyeksi[[k]]
  }

  list_hasil_fproyeksi = hasil_pilih_model

  return(list_hasil_fproyeksi)
}

fcetak = function(list_hasil_fproyeksi) {
  nama_wilayah = names(list_hasil_fproyeksi)

  # Ini nama tiap tabel data di tiap wilayah
  namadf = Map(names, list_hasil_fproyeksi)

  # Nama tiap tabel di tiap wilayah disambung dengan nama wilayah, lalu
  # setiap huruf awal katanya diubah ke huruf kapital.
  # Inspirasi kode dari Onyambu https://stackoverflow.com/a/52120292/14812170

  namadf_baru = Map(function(x, y) paste(x, y, sep = "_"), namadf, names(namadf))
  namadf_baru_format = lapply(namadf_baru, function(nama) {
    nomor = seq_len(length(namadf_baru[[1]]))
    gsub(pattern = "_", replacement = " ", x = nama) |>
      gsub("\\b(\\pL)", "\\U\\1", x = _, perl = TRUE)
  })
  nomor = seq_along(namadf_baru_format[[1]])
  for (k in seq_along(namadf_baru_format)) {
    namadf_baru_format[[k]] = Map(
      function(x, y) paste("Tabel", x, y),
      nomor, namadf_baru_format[[k]][nomor]
    )
  }

  # Nama tiap tabel itu dijadikan data frame supaya bisa dicetak di file Excel.
  ubah_jadi_tabel = function(x) {
    x |>
      as.data.table() |>
      setnames(" ")
  }
  tabel_namadf_baru_format = Map(
    function(x) Map(ubah_jadi_tabel, x),
    namadf_baru_format
  )

  dicetak = vector("list", length(nama_wilayah))
  names(dicetak) = nama_wilayah
  for (k in seq_along(nama_wilayah)) {
    dicetak[[k]] = vector("list", 2 * length(namadf_baru[[k]]))
    names(dicetak[[k]]) = rep(namadf_baru[[k]], each = 2)
  }

  posisi = seq_len(2 * length(namadf[[1]]))
  posisi_genap = posisi[posisi %% 2 == 0]
  posisi_gasal = posisi[posisi %% 2 == 1]
  for (k in seq_along(nama_wilayah)) {
    list_hasil_fproyeksi[[k]][["parameter_model_terbaik_semua_data"]] =
      list_hasil_fproyeksi[[k]][["parameter_model_terbaik_semua_data"]] |>
      summary() |>
      coef() |>
      as.data.table(keep.rownames = TRUE) |>
      setnames(c("Parameter", "Nilai", "Error Standar", "Nilai t", "Nilai p (p-value)"))
    dicetak[[k]][posisi_gasal] = Map(function(x, y) {
      dicetak[[k]][[x]] = tabel_namadf_baru_format[[k]][[y]]
    }, posisi_gasal, seq_along(tabel_namadf_baru_format[[k]]))
    dicetak[[k]][posisi_genap] = Map(function(x, y) {
      dicetak[[k]][[x]] = list_hasil_fproyeksi[[k]][[y]]
    }, posisi_genap, seq_along(list_hasil_fproyeksi[[k]]))
    colnames(dicetak[[k]][[2]]) =
      c("Tahun", "Urutan Tahun", "Penduduk (Jiwa)")
    colnames(dicetak[[k]][[4]]) = c(
      "Jenis Model", "Nama Parameter Model",
      "Nilai Parameter Model"
    )
    colnames(dicetak[[k]][[6]]) = c(
      "Tahun", "Penduduk (Jiwa)", "Urutan Tahun",
      "Prediksi Model Aritmatik",
      "Prediksi Model Eksponensial",
      "Prediksi Model Geometrik",
      "Prediksi Model Logistik",
      "Selisih Model Aritmatik",
      "Selisih Model Eksponensial",
      "Selisih Model Geometrik",
      "Selisih Model Logistik"
    )
    dicetak[[k]][[8]][,2:4] = apply(dicetak[[k]][[8]][,2:4], 2, round, digits = 2) |>
      as.data.table()
    colnames(dicetak[[k]][[8]]) = c(
      "Jenis Model",  "MAPE Mean Absolute Percentage Error (%)",
      "MAE Mean Absolute Error (Jiwa)",
      "RSS Residual Sum of Squares (Jiwa Kuadrat)"
    )
    colnames(dicetak[[k]][[10]]) = c("MAPE Terkecil", "MAE Terkecil", "RSS Terkecil")
    colnames(dicetak[[k]][[12]]) = "Jenis Model Terbaik"
    dicetak[[k]][[14]][,2:5] = apply(dicetak[[k]][[14]][,2:5], 2, round, digits = 3) |>
      as.data.table()
    colnames(dicetak[[k]][[16]]) =
      c(
        "Tahun", "Urutan Tahun", "Penduduk (Jiwa)",
        "Batas Bawah Interval Prediksi 95%", "Batas Bawah Interval Prediksi 95%"
      )
  }

  buku = createWorkbook()
  nbarisdf = vector("list", length(dicetak))
  for (k in seq_along(nama_wilayah)) {
    nbarisdf[[k]] = vapply(dicetak[[k]], nrow, numeric(1))
    addWorksheet(buku, sheetName = nama_wilayah[k])
    for (j in seq_along(dicetak[[k]])) {
      writeData(buku,
        sheet = nama_wilayah[k], borders = "all",
        borderColour = "orange3", borderStyle = "thin",
        headerStyle = createStyle(textDecoration = "bold"),
        x = dicetak[[k]][[j]],
        startRow = if (j < 2) {
          1
        } else {
          (cumsum(nbarisdf[[k]] + 2) - nbarisdf[[k]])[j]
        }
      )
    }
  }
  # saveWorkbook(buku, file = "hasil_model.xlsx", overwrite = TRUE)
  hasil = list(buku_cetak = buku, tabel_tampil = dicetak)
  return(hasil)
}

frumus = function(model, param_model, tahun_awal){
  param = param_model |> coef() |> round(4)
  r.aritmatik = function(m, c) {
    rumus = bquote(P(t) == .(m) (t -.(tahun_awal)) + .(c) ~Jiwa)
    return(rumus)
  }
  r.eksponensial = function(a, r) {
    rumus = bquote(P(t) == .(a)*e^.(r)(t - .(tahun_awal)) ~ Jiwa)
    return(rumus)
  }

  r.geometrik = function(a, r) {
    if(r >= 0) {
      rumus = bquote(P(t) == .(a) (1 + .(r))^(t-.(tahun_awal)) ~ Jiwa)
    }
    else {
      rumus = bquote(P(t) == .(a) (1 - .(-r))^(t-.(tahun_awal)) ~ Jiwa)
    }
    return(rumus)
  }

  r.logistik = function(k, a, r) {
    rumus = bquote(P(t) == .(k) / (1 + .(a) * e^{-.(r)(t-.(tahun_awal))}) ~ Jiwa)
    return(rumus)
  }

  fungsi.hasil =
    switch(model,
     "aritmatik" = r.aritmatik(m = unname(param["m_lin"]),
                               c = unname(param["c_lin"])),
     "eksponensial" = r.eksponensial(a = unname(param["a_eks"]),
                                     r = unname(param["r_eks"])),
      "geometrik" = r.geometrik(a = unname(param["a_geo"]),
                                r = unname(param["r_geo"])),
      "logistik" = r.logistik(k = param["k_log"], a = unname(param["a_log"]),
                                                r = unname(param["r_log"])))
  return(fungsi.hasil)
}
fkurva = function(x, model, param_model) {
  param = coef(param_model)

  f.aritmatik = function(x, m, c) {
    y = m * x + c
    return(y)
  }
  f.eksponensial = function(x, a, r) {
    y = a * exp(r * x)
    return(y)
  }
  f.geometrik = function(x, a, r) {
    y = a * (1 + r)^x
    return(y)
  }
  f.logistik = function(x, k, a, r) {
    y = k / (1 + a * exp(-r * x))
    return(y)
  }

  fungsi.hasil = switch(model,
    "aritmatik" = f.aritmatik(x, m = param["m_lin"], c = param["c_lin"]),
    "eksponensial" = f.eksponensial(x, a = param["a_eks"], r = param["r_eks"]),
    "geometrik" = f.geometrik(x, a = param["a_geo"], r = param["r_geo"]),
    "logistik" = f.logistik(x,
      k = param["k_log"], a = param["a_log"],
      r = param["r_log"]
    )
  )
  return(fungsi.hasil)
}


fgambar = function(hasil_fproyeksi, nama_wilayah, tahun_proyeksi = 2030:2045) {

  data_asli = hasil_fproyeksi[[nama_wilayah]][[1]]
  data_proyeksi = hasil_fproyeksi[[nama_wilayah]][[8]]
  n_data_asli = nrow(data_asli)
  n_data_proyeksi = nrow(data_proyeksi)
  data_gabungan = rbind(
    data_asli,
    data_proyeksi[, c("tahun", "penduduk", "urutan")]
  )
  data_gabungan$`Sumber Data` = c(
    rep("Data BPS", n_data_asli),
    rep("Data Prediksi Model", n_data_proyeksi)
  )

  data_kurva = data.frame(
    x_kurva = seq(min(data_gabungan$tahun), max(data_gabungan$tahun), by = 0.1)
  )
  data_kurva$y_kurva = fkurva(data_kurva$x_kurva - data_gabungan$tahun[1],
                              model = hasil_fproyeksi[[nama_wilayah]]$jenis_model_terbaik[[1]],
                              param_model = hasil_fproyeksi[[nama_wilayah]]$parameter_model_terbaik_semua_data
                             )
  modterbaik = hasil_fproyeksi[[nama_wilayah]]$jenis_model_terbaik[[1]]
  param_model_terbaik =
    hasil_fproyeksi[[nama_wilayah]]$parameter_model_terbaik_semua_data

  tahun_gabungan = as.integer(c(data_asli$tahun, data_proyeksi$tahun))
  penduduk_ribuan = mean(data_asli$penduduk) > 1e4 && mean(data_asli$penduduk) < 1e6
  penduduk_jutaan = mean(data_asli$penduduk) >= 1e6
  ggplot(data = data_asli, aes(x = as.integer(tahun))) +
    geom_ribbon(
      data = data_proyeksi, aes(
        x = as.integer(tahun), y = penduduk,
        ymin = batas_bawah, ymax = batas_atas,
      fill = "Interval Prediksi 95%"), alpha = 0.5
    ) +
    geom_line(
      data = data_kurva, aes(x = x_kurva, y = y_kurva), col = "tomato",
      lty = 1, lwd = 0.75
    ) +
    geom_point(
      data = data_gabungan, aes(
        y = penduduk,
        shape = `Sumber Data`
      ),
      size = 2
    ) +
    labs(
      title = paste("Hasil Pemodelan Data Penduduk", nama_wilayah),
      subtitle = frumus(model = hasil_fproyeksi[[nama_wilayah]]$jenis_model_terbaik[[1]],
                        param_model = hasil_fproyeksi[[nama_wilayah]]$parameter_model_terbaik_semua_data,
                        tahun_awal = data_gabungan$tahun[1]),
      x = "Tahun",
      y = if (isTRUE(penduduk_ribuan)){
        "Penduduk (ribu jiwa)"
      } else if (isTRUE(penduduk_jutaan)) {
        "Penduduk (juta jiwa)"
      } else "Penduduk (jiwa)") +
    theme_bw() +
    scale_x_continuous(breaks = seq(min(data_gabungan$tahun),
                                    max(data_gabungan$tahun), by = 5)) +
    scale_y_continuous(labels = function(y) {
      if (isTRUE(penduduk_ribuan)) {
        y / 1e3
      } else if (isTRUE(penduduk_jutaan)){
        y / 1e6
      } else y
    }) +
    scale_shape_manual(values = c(16, 10)) +
    scale_fill_manual(name = "", values = c("Interval Prediksi 95%" = "skyblue"))+
    theme(legend.position = "top",
          legend.title = element_blank(),
          theme(axis.text.x = element_text(size = 14),
                axis.text.y = element_text(size = 14),
                axis.title.x = element_text(size = 20),
                axis.title.y = element_text(size = 20)))
}
