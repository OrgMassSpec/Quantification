## This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.

library(shiny)

# Define UI for application that draws a histogram
ui <- fluidPage(
   
   # Application title
   titlePanel("Calibration Standard Selection"),
   
   # Sidebar with a slider input for number of bins 
   sidebarLayout(
      sidebarPanel(
         sliderInput("bins",
                     "Number of bins:",
                     min = 1,
                     max = 50,
                     value = 30)
      ),
      
      # Show a plot of the generated distribution
      mainPanel(
         plotOutput("distPlot")
      )
   ),
   
   # Copy the chunk below to make a group of checkboxes
   checkboxGroupInput("checkGroup", label = h3("Checkbox group"), 
                      choices = list("Choice 1" = 1, "Choice 2" = 2, "Choice 3" = 3),
                      selected = 1),
   
   
   hr(),
   fluidRow(column(3, verbatimTextOutput("value")))
)

# Define server logic required to draw a histogram
server <- function(input, output) {
   
   output$distPlot <- renderPlot({
      # generate bins based on input$bins from ui.R
      x    <- faithful[, 2] 
      bins <- seq(min(x), max(x), length.out = input$bins + 1)
      
      # draw the histogram with the specified number of bins
      hist(x, breaks = bins, col = 'darkgray', border = 'white')
   })
   
   # You can access the values of the widget (as a vector)
   # with input$checkGroup, e.g.
   output$value <- renderPrint({input$checkGroup})
   
}

# Run the application 
shinyApp(ui = ui, server = server)

