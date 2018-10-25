## This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.

library(shiny)
library(tidyverse)

x <- read_csv("TestData.csv")
IntStdConc <- 5 # ppb

calibration_data <- x %>% 
  filter(Type == "Cal") %>%
  mutate(calResponseRatio = TargetResponse / IntStdResponse,
         calAmountRatio = ExptConc / IntStdConc)

lm_result <- summary(lm(calibration_data$calResponseRatio ~ calibration_data$calAmountRatio))

slope <- signif(lm_result$coefficients[2,1], 3) 
intercept <- signif(lm_result$coefficients[1,1], 3)

sample_data <- x %>%
  filter(Type == "Sample") %>%
  mutate(calResponseRatio = TargetResponse / IntStdResponse,
         amount = (calResponseRatio - intercept) * (IntStdConc / slope))

ckbox <- paste0("'", calibration_data$SampleName, "'", "=", calibration_data$ExptConc, collapse = ",")

# Define UI for application that draws a histogram
ui <- fluidPage(
   
   # Application title
   titlePanel("Calibration Standard Selection"),
   
   # Sidebar with a slider input for number of bins 
   sidebarLayout(
      sidebarPanel(
         # sliderInput("bins",
         #             "Number of bins:",
         #             min = 1,
         #             max = 50,
         #             value = 30)
        checkboxGroupInput("checkGroup", label = h3("Calibration Levels (ng/mL)"), 
                           choices = c(calibration_data$ExptConc),
                           selected = c(calibration_data$ExptConc))
      ),
      
      # Show a plot of the generated distribution
      mainPanel(
         plotOutput("calibrationCurve")
      )
   ),
   
   # Copy the chunk below to make a group of checkboxes
   # checkboxGroupInput("checkGroup", label = h3("Calibration Levels"), 
   #                    choices = list("Choice 1" = 1, "Choice 2" = 2, "Choice 3" = 3),
   #                    selected = 1),
 
   hr(),
   fluidRow(column(3, verbatimTextOutput("value")))
)

# Define server logic required to draw a histogram
server <- function(input, output) {
   
   # output$distPlot <- renderPlot({
   #    # generate bins based on input$bins from ui.R
   #    x    <- faithful[, 2] 
   #    bins <- seq(min(x), max(x), length.out = input$bins + 1)
   #    
   #    # draw the histogram with the specified number of bins
   #    hist(x, breaks = bins, col = 'darkgray', border = 'white')
   #    })
   
   output$calibrationCurve <- renderPlot({
     
     calibration_curve_log <- ggplot() +
       geom_smooth(data = calibration_data, 
                   aes(calAmountRatio, calResponseRatio), 
                   method = lm, color = "grey", size = 1) +
       geom_point(data = calibration_data,
                  aes(calAmountRatio, calResponseRatio),
                  shape = 21, colour = "red", fill = NA, size = 3, stroke = 1.5) +
       geom_point(data = sample_data,
                  aes(x = amount/IntStdConc, y = calResponseRatio)) +
       xlim(NA, 50) +
       ylim(NA, 100) +
       geom_text(mapping = aes(calAmountRatio, calResponseRatio, label = ExptConc),
                 data = calibration_data,
                 size = 4, color = "red", vjust = -0.5, hjust = 0, nudge_x = -0.5) +
       labs(x = 'Amount Ratio (ng/mL)', y = 'Response Ratio') +
       labs(title = 'Selection Calibration Curve, Nicotine') +
       theme_bw()
     print(calibration_curve_log)
   })
   
   # You can access the values of the widget (as a vector)
   # with input$checkGroup, e.g.
   output$value <- renderPrint({input$checkGroup})
   
}

# Run the application 
shinyApp(ui = ui, server = server)

