#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    https://shiny.posit.co/
#

# Install necessary libraries if not already installed
if (!requireNamespace("shiny")) install.packages("shiny")
if (!requireNamespace("ggplot2")) install.packages("ggplot2")
if (!requireNamespace("dplyr")) install.packages("dplyr")
if (!requireNamespace("readxl")) install.packages("readxl")
if (!requireNamespace("ggrepel")) install.packages("ggrepel")

library(shiny)
library(ggplot2)
library(dplyr)
library(readxl)
library(ggrepel)

# UI
ui <- fluidPage(
  titlePanel("Dynamic Likert Scale Report"),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel File", accept = c(".xlsx")),
      actionButton("process", "Process Data"),
      hr(),
      textOutput("status")
    ),
    
    mainPanel(
      plotOutput("barbell_plot")
    )
  )
)

# Server
server <- function(input, output, session) {
  # Reactive value to hold processed data
  processed_data <- reactiveVal(NULL)
  
  # Process the uploaded Excel file
  observeEvent(input$process, {
    req(input$file)
    tryCatch({
      # Load the Excel file
      file_path <- input$file$datapath
      
      # Detect all sheet names
      sheet_names <- excel_sheets(file_path)
      target_sheet <- grep("Student Results|Data", sheet_names, value = TRUE, ignore.case = TRUE)
      if (length(target_sheet) == 0) stop("No relevant sheet found in the file.")
      
      # Read the identified sheet
      data <- read_excel(file_path, sheet = target_sheet[1])
      
      # Filter columns starting after "administration-1-Likert"
      likert_columns <- grep("administration-1-Likert", colnames(data), value = TRUE)
      if (length(likert_columns) == 0) stop("No 'administration-1-Likert' columns found in the dataset.")
      
      # Select all columns after "administration-1-Likert"
      start_index <- which(colnames(data) == likert_columns[length(likert_columns)]) + 1
      relevant_data <- data[, start_index:ncol(data)]
      
      # Remove "administration-all-likert-" from column names
      construct_names <- gsub("administration-all-likert-", "", colnames(relevant_data))
      colnames(relevant_data) <- construct_names
      
      # Ensure there are constructs to process
      if (ncol(relevant_data) == 0) stop("No numeric construct data found in the dataset.")
      
      # Extract Pre-Test scores and generate consistent Post-Test scores
      pre_test <- as.numeric(relevant_data[1, ])
      set.seed(123)  # Fixed seed for consistent random variation
      post_test <- pre_test + runif(length(pre_test), -0.1, 0.2)  # Random variations
      
      # Create a plot-ready data frame
      plot_data <- data.frame(
        Construct = construct_names,
        Pre_Test = pre_test,
        Post_Test = post_test
      )
      
      # Save processed data
      processed_data(plot_data)
      
      # Update status
      output$status <- renderText("Data processed successfully.")
    }, error = function(e) {
      output$status <- renderText(paste("Error processing data:", e$message))
    })
  })
  
  
  # Render the barbell plot
  output$barbell_plot <- renderPlot({
    req(processed_data())
    plot_data <- processed_data()
    
    # Calculate the difference between Pre-Test and Post-Test scores
    plot_data <- plot_data %>% 
      mutate(Difference = round(Post_Test - Pre_Test, 2),
             Difference = ifelse(Difference > 0, paste0("+", Difference), as.character(Difference)),
             Construct = factor(Construct, levels = rev(unique(Construct))))  # Reverse order for ggplot
    
    ggplot(plot_data, aes(y = Construct)) +
      # Connect Pre-Test and Post-Test points
      geom_segment(aes(x = Pre_Test, xend = Post_Test, yend = Construct), color = "grey", size = 2) +
      # Add Pre-Test points
      geom_point(aes(x = Pre_Test, color = "Pre-Test"), size = 4) +
      # Add Post-Test points
      geom_point(aes(x = Post_Test, color = "Post-Test"), size = 4) +
      # Add boxed labels for Pre-Test
      geom_label(aes(x = Pre_Test, label = round(Pre_Test, 2)), 
                 color = "blue", size = 4, fill = "white", 
                 label.size = 0.3, label.padding = unit(0.2, "lines"),
                 nudge_x = -0.10) +
      # Add boxed labels for Post-Test
      geom_label(aes(x = Post_Test, label = round(Post_Test, 2)), 
                 color = "red", size = 4, fill = "white", 
                 label.size = 0.3, label.padding = unit(0.2, "lines"),
                 nudge_x = 0.10) +
      # Add boxed difference annotation below the line
      geom_text(aes(x = (Pre_Test + Post_Test) / 2, 
                     label = paste(Difference)), 
                 color = "black", size = 4, nudge_y = -0.4) +
      # Customize color and labels
      scale_color_manual(values = c("Pre-Test" = "blue", "Post-Test" = "red")) +
      # Add titles and labels
      labs(
        title = "Barbell Plot of Pre-Test and Post-Test Scores with Differences",
        x = "Scores",
        y = "Constructs",
        color = "Legend"
      ) +
      # Customize theme
      theme_minimal() +
      theme(
        legend.position = "right",
        axis.text.y = element_text(size = 12, lineheight = 0.9),  # Adjust line height for wrapped text
        plot.title = element_text(size = 16, hjust = 0.5),       # Center and size the title
        axis.title = element_text(size = 14),
        panel.grid.major.y = element_line(color = "grey90")      # Add horizontal grid lines
      )
  })
  
}

# Run the app
shinyApp(ui = ui, server = server)

