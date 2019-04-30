#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
library(interactexcel)
library(readxl)

# Define UI for application that draws a histogram
ui <- fluidPage(
  titlePanel("Run Excel Load Cases"),
  sidebarLayout(
    sidebarPanel(
      fileInput('file1', 'Choose xls input file',
                accept = c(".xls")
      ),
      fileInput('file2', 'Choose xls calculation spreadsheet',
                accept = c(".xls")),
      textInput('sheet_name', 'Specify Sheet Name')
      
    ),
    mainPanel(
      tabsetPanel(
        tabPanel('Input', tableOutput('contents')),
        tabPanel('Output')
      )
    )
  )
)

# Define server logic required to draw a histogram
server <- function(input, output) {
  output$contents <- renderTable({
    
    req(input$file1)
    
    inFile <- input$file1
    
    run_matrix <- read_excel(inFile$datapath, 1)
    
    file_location <- input$file2 
    sheet <- input$sheet_name
    
    solved <- 
      use_run_matrix(
        run_matrix, 
        file_location,
        sheet)
    
    solved
  })
}

# Run the application 
shinyApp(ui = ui, server = server)

