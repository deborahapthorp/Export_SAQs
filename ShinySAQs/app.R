# Adapted from the R Bloggers tutorial, How to Generate Word Docs in Shiny with officer, flextable, and shinyglide, 
# at https://www.r-bloggers.com/2022/11/how-to-generate-word-docs-in-shiny-with-officer-flextable-and-shinyglide/
# Load libraries
library(DT)
library(officer)
library(rmarkdown)
library(shiny)
library(shinyglide)
library(shinyWidgets)
library(readxl)
library(tidyverse) # for tidy code 
library(janitor) # for cleaning variable names 
library(tools)

# Define UI
ui <- fluidPage(
  # Add application title
  titlePanel("Export Moodle Short Answer Questions to a Word Document"),
  # Add glide functionality
  glide(
    height = "100%",
    controls_position = "top",
    
    
    screen(
      h1("Instructions 1 - Download Moodle quiz responses"), 
      
      mainPanel(
        h4("Download the responses from your Moodle quiz by navigating to the quiz and clicking on the cog at the top right."),
        h4("Select the 'responses' option."),
        img(src='quiz_responses_1.png', align = "center"),
        width = 12
      )
      ),
    
    screen(
      h1("Instructions 2 - Download Moodle quiz responses"), 
      mainPanel(
        h4("In the 'Display Options' section, tick only 'response'."), 
        HTML("<br></br>"),
        img(src='Display_options.png', align = "center"),
        
        HTML("<br></br>"),
        
        h4("You can download either as an Excel file or as a csv - either is fine."), 
        img(src='download_as_Excel.png', align = "left"),
        
        HTML("<br></br>"),
  
        width = 12
      )
      
    ),
    
    screen(
      h1("Instructions 3 - What to do next"), 
      mainPanel(
        h4("Your data should look something like this:"),
        img(src='data_example.png', align = "center"),
        HTML("<br></br>"),
        h4("You may want to remove the multiple choice responses and non-essential columns now, but it is not required."),
        h4("On the next screen, you will be able upload these responses to the app and export them as a Word document."),
        
        width = 12
      )
      
    ),

    screen(
      # Contents of screen 2
      h1("Upload Data"),
      sidebarLayout(
        sidebarPanel(
          # Upload data
          fileInput(
            inputId = "raw_data",
            label = "Upload .csv or .xlsx",
            accept = c(".csv", ".xlsx")
          )
        ),
        mainPanel(
          # Preview data
          DTOutput("preview_data")
        )
      ),
      # Disable Next button until data is uploaded
      next_condition = "output.preview_data !== undefined"
    ),
    screen(
      # Contents of screen 3
      h1("Create Word Document"),
      sidebarLayout(
        sidebarPanel(
          # Set word document title
          textInput(
            inputId = "document_title",
            label = "Set Word Document Title",
            value = "My Title"
          ),
          # Select grouping variable
          pickerInput(
            inputId = "student_id",
            label = "Select Student ID column",
            choices = NULL,
            multiple = FALSE
          ),
          # Select the variables used to compare groups
          pickerInput(
            inputId = "SAQs",
            label = "Select all short answer questions",
            choices = NULL,
            multiple = TRUE
          ),
          # Set word document filename
          textInput(
            inputId = "filename",
            label = "Set Word Document Filename",
            value = "my_SAQs"
          ),
          
          # Download word document
          downloadButton(
            outputId = "download_word_document", 
            label = "Download Word Document"
          )
        ),
        mainPanel(
         #  Preview results
           htmlOutput("result")
        )
      )
    )
  )
)
# Define server
server <- function(input, output, session) {
  # Handle input data
  data <- reactive({
    req(input$raw_data)
    ext <- tools::file_ext(input$raw_data$name)
    switch(ext,
           csv = readr::read_csv(input$raw_data$datapath),
           xlsx = readxl::read_excel(input$raw_data$datapath),
           validate("Invalid file. Please upload a .csv or .xlsx file")
    )
  })
  
  # Create DataTable that shows uploaded data
  output$preview_data <- DT::renderDT({ data() |> 
      DT::datatable(
        rownames = FALSE,
        options = list(
          searching = FALSE,
          filter = FALSE
        )
      )
  })
  
  # Update input choices when data is uploaded
  observeEvent(data(), {
    var_names = colnames(data())
    
    updatePickerInput(
      session = session,
      inputId = "student_id",
      choices = var_names
    )
    
    updatePickerInput(
      session = session,
      inputId = "SAQs",
      choices = var_names
    )
  })

  
  varlistQ <- reactive({input$SAQs})
  nQuestions <- reactive({length(varlistQ())})
  all_responses <- reactive({data()})
  rows <- reactive({nrow(all_responses())})
  student <- reactive({input$student_id})
  
  # Download Handler
  output$download_word_document <- shiny::downloadHandler(
    filename = function() {paste(input$filename, ".docx", sep = "")},
    content = function(file) {
      doc1 <- read_docx() # calls function to create a document using officeR
      # Loop that makes a separate section for each question
      for (i in 1:nQuestions()) {
        qName <- varlistQ()[i] # Question name
        doc1 <-
          body_add_par(doc1, value = qName, style = "heading 1") # Add question name at top
        for (q in 1:rows()) {
          # student loop (could also make this anonymous by using student no.)
          doc1 <- body_add_par(doc1,
                               value = all_responses()[[student()]][q] ,
                               style = "heading 2",
                               pos = "after" ) # Add student name to doc
          doc1 <- body_add_par(doc1,
                               value = all_responses()[[varlistQ()[i]]][q],
                               style = "Normal",
                               pos = "after") # add response
        }
      }
      
      doc1 |>
        print(target = file)
    }
  )
}

# Run app
shinyApp(ui, server)