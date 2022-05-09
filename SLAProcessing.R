#https://www.r-bloggers.com/2019/07/excel-report-generation-with-shiny/

library("ERICDataProc")
SLA_config <- setup_SLA_config_values()

options(shiny.maxRequestSize=400*1024^2)

  # minimal Shiny UI
ui <- fluidPage(
  titlePanel("SLA processing"),
  tags$br(),

  fileInput("csvfile", "Choose CSV File",
            accept = c(
              "text/csv",
              "text/comma-separated-values,text/plain",
              ".csv")
  ),

  radioButtons(inputId = "slatype",label="SLA Type",c("Local authority"="la","NNPA/AONB"="nnpa")),

  checkboxInput(inputId = "datasource",label = "Data in Recorder format?",TRUE),

  textInput(inputId = "outputfile",label = "Output filename",value = ""),
  textOutput(outputId = "msg"),



  downloadButton(
    outputId = "okBtn",
    label = "Process SLA data")


)

# minimal Shiny server
server <- function(input, output) {


  output$okBtn <- downloadHandler(
    filename = function() {
      ifelse(stringr::str_ends(input$outputfile,'xlsx'),input$outputfile,paste0(input$outputfile,'.xlsx'))

    },
    content = function(file) {

      SLA_type <- ifelse((input$slatype=='la'),1,0)
      recorder <- ifelse(input$datasource,1,0)


      inFile <- input$csvfile

      if (is.null(inFile))
        return(NULL)

      raw_data <- read.csv(inFile$datapath,header = TRUE)


      if (SLA_type==1) {

        OutputCols <- SLA_config["LAOutputCols"]
        newColNames <- SLA_config["LAColNames"]
        locationCol <- SLA_config["LAlocationCol"]
        abundanceCol <- SLA_config["LAabundanceCol"]
        commentCol <- SLA_config["LAcommentCol"]
        recorderCol <- SLA_config["LArecorderCol"]
        lastCol <- SLA_config["LAlastCol"]
      } else  {
        OutputCols <- SLA_config["SLAOutputCols"]
        newColNames <- SLA_config["SLAColNames"]
      }



      started <- Sys.time()

      #Create output workbook
      XL_wb <- openxlsx::createWorkbook()

      inFile <- input$csvfile

      if (is.null(inFile))
        return(NULL)

      raw_data <- read.csv(inFile$datapath,header = TRUE)

        if (recorder==1) {
          #If from Recorder change some column names
          raw_data <- change_recorder_col_names(raw_data)
        }

        #Format the date
        raw_data$Sample.Dat <- formatDates(raw_data$Sample.Dat)

        #Filter out the records we don't want
        raw_data <- filter_by_survey(raw_data)

        #Remove GCN licence zero counts
        raw_data <- filter_GCN(raw_data)

        #Merge location columns (EA doesnt need this)
        raw_data <- merge_location_cols(raw_data)

        if (SLA_type!=1) {
          #Get broad group (NNPA & North Pennines only)
          raw_data <-  dplyr::left_join(raw_data,ERICDataProc:::group_LU, by=c("Taxon.grou"="group"))
        }

        #Fish
        raw_data <- do_data_fixes(raw_data)

        #Set protected flag
        # #Not sure prot species is correct - = not "in"

        raw_data <- set_protected_flag(raw_data)


        # #Sort out designations
        raw_data <- fix_designations(raw_data,TRUE)

        # Easting & northing calculations
        raw_data <- E_and_N_calcs(raw_data)

        if (SLA_type==1) {
          #Put an empty column on the end
          raw_data$Empty1 <- ''
          raw_data$Empty2 <- ''
          raw_data$Empty3 <- ''
        }

        #Get the columns we're going to output
        #AED this needs checking

        outputdata <- format_and_check_SLA_data(raw_data,OutputCols,newColNames)

        sheet_name = 'SLA data'
        XL_wb <- openxlsx::createWorkbook()


        XL_wb <- format_SLA_Excel_output(XL_wb,sheet_name,outputdata,SLA_config,SLA_type)


        openxlsx::saveWorkbook(XL_wb,file,overwrite = TRUE)
    }
  )


}

shinyApp(ui, server)
