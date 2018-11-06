library(shiny)
library(tidyverse)

x <- read_csv("TestData.csv")
IntStdConc <- 5 # ppb

calibration_data <- x %>% 
  filter(Type == "Cal") %>%
  mutate(calResponseRatio = TargetResponse / IntStdResponse,
         calAmountRatio = ExptConc / IntStdConc)

sample_data <- x %>%
  filter(Type == "Sample") %>%
  mutate(calResponseRatio = TargetResponse / IntStdResponse)

ckbox <- paste0("'", calibration_data$SampleName, "'", "=", calibration_data$ExptConc, collapse = ",")

# Define UI for application that draws a histogram
ui <- fluidPage(
   
   # Application title
   titlePanel("Calibration Standard Selection"),
   
   # Sidebar with a slider input for number of bins 
   sidebarLayout(
      sidebarPanel(
        checkboxGroupInput("checkGroup", label = h3("Calibration Levels (ng/mL)"), 
                           choices = c(calibration_data$ExptConc),
                           selected = c(calibration_data$ExptConc))
      ),
      
      # Show a plot of the generated distribution
      mainPanel(
         plotOutput("calibrationCurve")
      )
   ),
   
   hr(),
   fluidRow(column(3, verbatimTextOutput("value")))
)

# Define server logic required to draw a histogram
server <- function(input, output) {
  
   output$calibrationCurve <- renderPlot({
     
     # Filter calibration curve by selection
     
     calibration_data <- calibration_data %>% 
       filter(ExptConc %in% input$checkGroup)
     
     lm_result <- summary(lm(calibration_data$calResponseRatio ~ calibration_data$calAmountRatio))
     slope <- signif(lm_result$coefficients[2,1], 3) 
     intercept <- signif(lm_result$coefficients[1,1], 3)
     
     data_range <- range(calibration_data$calResponseRatio)

     sample_data <- sample_data %>%
       filter(calResponseRatio < data_range[2], calResponseRatio > data_range[1]) %>%
       mutate(amount = (calResponseRatio - intercept) * (IntStdConc / slope))
     
     # TODO Show data frame of sample points within current calibration point range
     
     calibration_curve <- ggplot() +
       geom_smooth(data = calibration_data,
                   aes(calAmountRatio, calResponseRatio),
                   method = lm, color = "grey", size = 1) +
       geom_point(data = calibration_data,
                  aes(calAmountRatio, calResponseRatio),
                  shape = 21, colour = "red", fill = NA, size = 3, stroke = 1.5) +
       geom_point(data = sample_data,
                  aes(x = amount/IntStdConc, y = calResponseRatio)) +
       geom_text(mapping = aes(calAmountRatio, calResponseRatio, label = ExptConc),
                 data = calibration_data,
                 size = 4, color = "red", vjust = -0.5, hjust = 0, nudge_x = -0.5) +
       labs(x = 'Amount Ratio (ng/mL)', y = 'Response Ratio') +
       #labs(title = 'Selection Calibration Curve, Nicotine') +
       labs(title = 'Calibration Curve, Nicotine (Target/IS)', 
            subtitle = paste('n=', nrow(calibration_data), 
                             ', slope=', signif(lm_result$coefficients[2,1], 3), 
                             ', intercept=', signif(lm_result$coefficients[1,1], 3),
                             ', multiple R-squared=', signif(lm_result$r.squared, 3), sep = '')) +
       theme_bw()
      
     
     print(calibration_curve)
   })
   
   observeEvent(input$calibrationCurve_dblclick, {
     brush <- input$calibrationCurve_brush
     if (!is.null(brush)) {
       ranges$x <- c(brush$xmin, brush$xmax)
       ranges$y <- c(brush$ymin, brush$ymax)
       
     } else {
       ranges$x <- NULL
       ranges$y <- NULL
     }
   })
   
   # You can access the values of the widget (as a vector)
   # with input$checkGroup, e.g.
   output$value <- renderPrint({input$checkGroup})
   
}

# Run the application 
shinyApp(ui = ui, server = server)

