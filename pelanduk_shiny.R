source("pelanduk_fungsi.R")
source("cek_kondisi.R")

fpaket(
  "shiny", "bslib", "openxlsx", "data.table", "minpack.lm",
  "ggplot2", "investr"
)


antarmuka <- page_fillable(
  HTML("<h2>Pelanduk: Aplikasi Pemodelan Data Penduduk (versi 0.1)</h2>"),
  theme = bs_theme(version = 5, preset = "minty"),
  layout_sidebar(
    border_radius = TRUE,
    sidebar = accordion(
      id = "samping", multiple = TRUE, open = TRUE,
      accordion_panel(
        "Unggah File",
        card(
          fileInput("file_data_penduduk_unggah",
                    HTML("File Data Penduduk(<em>.xlsx</em>)"),
                    accept = ".xlsx",
                    width = "100%",
                    buttonLabel = "Unggah",
                    placeholder = "..."
          ),
          card(actionButton("ok_unggah", "Tampilkan Data", class = 'btn-info')),
          card(actionButton("ok_jalankan", "Jalankan Model", class = 'btn-success',
                            disabled = TRUE))
        )
      ),
      accordion_panel(
        "Pengaturan",
        card(
          sliderInput("tahun_proyeksi", "Tahun Proyeksi",
                      min = 1900, max = 2100, value = c(2024, 2030),
                      sep = ""
          )
        ),
        card(
          sliderInput("persen_uji", HTML("Persentase Data Uji (%)"),
                      value = 20, min = 1, max = 100
          )
        )
      ),
      accordion_panel("Unduh Hasil", value = "tombol_unduh",
                      HTML("<em>Hasil model belum tersedia.</em>"))
    ),

    navset_card_pill(
      card_header(
        selectInput("wilayah_tabel_data", "Nama Wilayah",
                    choices = NULL, selected = NULL),
      ),
      nav_panel(
        "Tabel dan Grafik Data Penduduk",

        layout_column_wrap(fill = TRUE, width = 1/2,
                           card(
                             htmlOutput("judul_tabel_data_penduduk"),
                             tableOutput("tabel_data_penduduk")),
                           card(full_screen = TRUE,
                                htmlOutput("judul_grafik_data_penduduk"),
                                card_body(
                                  class = "p-0",
                                  plotOutput("grafik_model"))

                           )
        )),
      nav_panel(
        "Informasi Model",
        layout_column_wrap(
          card(
            tags$b("Parameter Semua Model dengan Data Latih"),
            tableOutput("parameter_model_latih")
          ),
          card(
            tags$b("Akurasi Semua Model dengan Data Uji"),
            tableOutput("akurasi_model")
          ),
          card(
            tags$b("Parameter Model Terbaik dengan Semua Data"),
            tableOutput("parameter_model_terbaik_semua_data")
          ),
          card(
            tags$b("Proyeksi Penduduk dengan Model Terbaik"),
            tableOutput("proyeksi_penduduk")
          )
        )
      )
    )
  )
)



pengatur <- function(input, output, session) {

  data_penduduk_unggah <- reactive({
    req(isolate(input$file_data_penduduk_unggah))
    nama_file <- isolate(input$file_data_penduduk_unggah$datapath)
    fbaca(nama_file)
  })

  observeEvent(input$ok_unggah, {
    updateActionButton(session, "ok_jalankan", "Jalankan Model", disabled = FALSE)
    updateSelectInput(session, "wilayah_tabel_data", "Nama Wilayah",
                      choices = names(data_penduduk_unggah()),
                      selected = names(data_penduduk_unggah())[1])
    output$judul_tabel_data_penduduk = renderText(
      paste("<b>Tabel Data Penduduk", input$wilayah_tabel_data,"</b"))
    output$tabel_data_penduduk <- renderTable(
      data_penduduk_unggah()[[input$wilayah_tabel_data]], width  = '40%',
      striped = TRUE, bordered = TRUE)
  })

  proyeksi_semua <- reactive(
    data_penduduk_unggah() |>
      fmodel(persen_tes = input$persen_uji) |>
      fproyeksi(tahun_proyeksi = input$tahun_proyeksi[1]:input$tahun_proyeksi[2])
  ) |>
    bindCache(input$persen_uji, input$tahun_proyeksi) |>
    bindEvent(input$persen_uji, input$tahun_proyeksi)

  proyeksi_wilayah <- reactive({
    proyeksi_semua()[input$wilayah_tabel_data]
  }) |>
    bindCache(input$wilayah_tabel_data, input$persen_uji, input$tahun_proyeksi) |>
    bindEvent(input$wilayah_tabel_data, input$persen_uji, input$tahun_proyeksi)

  observeEvent(input$ok_jalankan, {

    output$judul_grafik_data_penduduk = renderText(paste("<b>Grafik Data Penduduk",
                                                         input$wilayah_tabel_data,"</b"))

    output$grafik_model <- renderPlot(
      proyeksi_wilayah() |> fgambar(nama_wilayah = input$wilayah_tabel_data)
    )

    output$tabel_param_model_latih <-
      renderTable(
        proyeksi_semua()[[input$wilayah_tabel_data]][["tabel_parameter"]],
        striped = TRUE, bordered = TRUE
      )
    output$parameter_model_latih <- renderTable(
      tabel_hasil_model()[[4]], striped = TRUE, bordered = TRUE
    )
    output$akurasi_model <- renderTable(
      tabel_hasil_model()[[8]], striped = TRUE, bordered = TRUE
    )
    output$parameter_model_terbaik_semua_data <- renderTable(
      tabel_hasil_model()[[14]], striped = TRUE, bordered = TRUE
    )
    output$proyeksi_penduduk <- renderTable(
      tabel_hasil_model()[[16]], striped = TRUE, bordered = TRUE
    )
    accordion_panel_update(id = "samping", target = "tombol_unduh",
                           card(id = "kartu_unduh_tabel",
                                downloadButton("unduh_file_tabel", "Unduh Tabel(.xlsx)",
                                               icon = icon("download")
                                )),
                           card(id = "kartu_unduh_gambar",
                                downloadButton("unduh_file_gambar", "Unduh Gambar(.png)",
                                               icon = icon("download")
                                ))
    )

  }, ignoreNULL = TRUE)


  tabel_hasil_model <- reactive(
    fcetak(
      proyeksi_semua()
    )[["tabel_tampil"]][[input$wilayah_tabel_data]]
  ) |>
    bindCache(input$wilayah_tabel_data, input$persen_uji, input$tahun_proyeksi)

  output$unduh_file_tabel <- downloadHandler(
    filename = function() {
      paste0(
        "Tabel Hasil Pemodelan Data Penduduk ",
        tools::file_path_sans_ext(input$file_data_penduduk_unggah$name),
        ".xlsx"
      )
    },
    content = function(file_hasil) {
      showModal(modalDialog(paste0("Mohon tunggu, sedang diunduh: ",
                                   paste0(
                                     "Tabel Hasil Pemodelan Data Penduduk ",
                                     tools::file_path_sans_ext(input$file_data_penduduk_unggah$name),
                                     ".xlsx"
                                   )),  easyClose = TRUE, size = "l"))
      fcetak(proyeksi_semua())[["buku_cetak"]] |>
        saveWorkbook(file = file_hasil, overwrite = TRUE)
      removeModal()
    }
  )


  output$unduh_file_gambar <- downloadHandler(
    filename = "",
    content = function(file_hasil) {
      names(proyeksi_semua()) |> lapply(function(x) {
        gambar_semua <- fgambar(proyeksi_semua(),
                                nama_wilayah = x,
                                tahun_proyeksi = input$tahun_proyeksi[1]:input$tahun_proyeksi[2]
        )
        png(paste0("Grafik Hasil Pemodelan Data Penduduk ", x, ".png"),
            width = 20, height = 10, units = "cm", res = 300)
        print(gambar_semua)
        dev.off()

        showModal(modalDialog(paste0("Sudah Mengunduh:
                                     Grafik Hasil Pemodelan Data Penduduk ", x, ".png"),
                              footer =  paste("Lokasi file:", getwd()),
                              easyClose = TRUE, size = "l"))
        Sys.sleep(1)
      })
      Sys.sleep(2)
      removeModal()
    }
  )
}

shinyApp(ui = antarmuka, server = pengatur)
