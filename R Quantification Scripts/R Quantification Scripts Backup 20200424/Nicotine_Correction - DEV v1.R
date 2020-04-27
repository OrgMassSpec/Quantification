#' ---
#' title: "Project: DEVELOPMENT VERSION | Matrix: | Analyte: Nicotine"
#' author: "Nate Dodder"
#' output:
#'  html_document:
#'    toc: true
#'    toc_float: 
#'      collapsed: false 
#'      smooth_scroll: false
#'    df_print: default
#'    mathjax: null
#' ---

#+ echo=FALSE
# Setup
    library(tidyverse)
    library(readxl)
    library(knitr)
    knitr::opts_chunk$set(echo = FALSE)
    knitr::opts_chunk$set(comment = NA)
    options("width" = 200)
    options(scipen = 999)

# source('Nicotine_Correction.R')
# rmarkdown::render('Nicotine_Correction.R')

# Checklist:
# 1. Enter project in title
# 2. Enter matrix in title
# 3. Enter instrument sequence name
# 4. Enter input file name
# 5. Enter/check internal standard spike concentration and extraction volume
# 6. Change script name in source() and rmarkdown::render().
# 7. See OneNote quantification checklist.

#' Quantification quality control for nicotine LC/MS/MS measurements.
#'
#' __Dilution:__ none. 

# Input variables

    # input (.xlsx) and generated output (.csv) file name
    file_name   = 'Data - Nicotine DEV'    

    intstd_conc = 5     # internal standard concentration (ng/mL)
    extract_vol = 3     # acetonitrile extraction volume (mL)

#+ results='asis'
cat('__Input data file name:__ ', file_name, '.xlsx. and __Reported as:__ ', file_name, '.csv.', sep = '')
#'

#+ results='asis'
cat('__Internal standard final concentration:__', intstd_conc, 'ng/mL.')
#'

#+ results='asis'
cat('__Extraction volume:__', extract_vol, 'mL.')
#'

#+ results='asis'
cat('__Results reported as:__ nicotine concentration (ng/mL) x ', extract_vol, ' mL = *ng nicotine*.', sep = '')
#'

#' # Uncorrected calibration 

    column_names <- c("SampleID", "DataFile", "Type", "Level", "DateTime", "AcqMethodFile", "DataAnalysisMethodFile",
                      "Comment", "TC_Conc", "Analyte_RT", "TC_Response", "Analyte_ManualIntegration", "Analyte_CalcConc",
                      "Analyte_FinalConc", "Analyte_Accuracy", "IS_ReponseRatio", "Analyte_Transition1_Ratio",
                      "Analyte_Transition1_ManualIntegration", "Analyte_Transition2_Ratio",
                      "Analyte_Transition2_ManualIntegration", "IS_RT", "IS_Response", "IS_Transition1_Ratio",
                      "IS_Transition1_ManualIntegration", "IS_Transition2_Ratio", "IS_Transition2_ManualIntegration")

    x <- read_excel(str_c(file_name, '.xlsx'), sheet = 'Sheet1', col_names = column_names, skip = 2)

    df_cal <-   x %>%
                filter(Type == 'Cal') %>%
                select(SampleID, TC_Response, IS_Response, TC_Conc) %>%
                mutate(Response_Ratio = TC_Response / IS_Response) %>%
                mutate(Conc_Ratio = TC_Conc / intstd_conc)

    df_spl <-   x %>%
                filter(Type == 'Sample') %>%
                select(SampleID, TC_Response, IS_Response) %>%
                mutate(Response_Ratio = TC_Response / IS_Response)

    df_spl_order <- x %>%
                    filter(Type == 'Sample') %>%
                    select(SampleID) %>%
                    mutate(Order = 1:nrow(df_spl))

    df_spl_masshunter <- x %>%
                filter(Type == 'Sample') %>%
                select(SampleID, Analyte_CalcConc) %>%
                rename(TC_Conc_MassHunter = Analyte_CalcConc)

    # Linear model of calibration curve
    fit <- lm(Response_Ratio ~ Conc_Ratio, data = df_cal)
    intercept   <- coefficients(fit)[1]
    slope       <- coefficients(fit)[2]

    # Calibration accuracy
    df_cal <-   df_cal %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept) / slope) * intstd_conc) %>%
                mutate(Accuracy_Pct = abs((((Measured_TC_Conc - TC_Conc) / TC_Conc) * 100) - 100))
    r_squared           <- summary(fit)$r.squared 
    mean_is_response    <- mean(df_cal$IS_Response)

#' ## Table: Uncorrected accuracy
#' If accuracy is within 90% to 110% for all calibration points, correction is not required. 

    df_cal_print <- df_cal %>%
                    select(-SampleID) %>%
                    select(TC_Conc, everything())

    kable(df_cal_print, digits = c(1, 0, 0, 2, 2, 1, 0))

# Sample concentration
    df_spl <-   df_spl %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept) / slope) * intstd_conc) %>%
                mutate(Measured_Conc_Ratio = Measured_TC_Conc / intstd_conc) %>%
                mutate(IS_Recovery = (IS_Response / mean_is_response) * 100)

    mean_is_recovery <- mean(df_spl$IS_Recovery)
    mean_tc_conc <- mean(df_spl$Measured_TC_Conc)

    # Check for zero responses and set outputs to zero
    df_spl <-   df_spl %>%
                mutate(Measured_TC_Conc = replace(Measured_TC_Conc, TC_Response == 0, 0)) %>%
                mutate(Measured_Conc_Ratio = replace(Measured_Conc_Ratio, TC_Response == 0, 0))

#' ## Plot: Uncorrected calibration

#+ results='asis'
cat('R^2^ = ', signif(r_squared, 4))
#'

#+ fig.width = 7, fig.height = 5

    df_cal_tmp <-   df_cal %>%
                    select(Conc_Ratio, Response_Ratio) %>%
                    mutate(Category = 'Calibration')

    df_spl_tmp <-   df_spl %>%
                    select(Measured_Conc_Ratio, Response_Ratio) %>%
                    rename(Conc_Ratio = Measured_Conc_Ratio) %>%
                    mutate(Category = 'Sample')

    df_uncorrected <- bind_rows(df_spl_tmp, df_cal_tmp)

    p <-    ggplot(df_uncorrected) + 
            geom_point(aes(Conc_Ratio, Response_Ratio, 
                color = factor(Category), shape = factor(Category), size = factor(Category))) +
            geom_abline(intercept = intercept, slope = slope, color = 'blue') +
            theme_bw() +
            scale_color_manual(values = c('blue', 'red')) +
            scale_shape_manual(values = c(16,17)) +
            scale_size_manual(values = c(3,4)) +
            labs(x = 'TC/IS concentration ratio', y = 'TC/IS response ratio')
    p

#' ## Plot: Uncorrected calibration low end
#+ fig.width = 4, fig.height = 4

    df_cal_tmp <- slice(df_cal_tmp, 1:4)

    p <-        ggplot(df_cal_tmp) + 
                geom_point(aes(Conc_Ratio, Response_Ratio), color = 'blue', shape = 16, size = 3) +
                geom_abline(intercept = intercept, slope = slope, color = 'blue') +
                theme_bw() +
                labs(x = 'TC/IS concentration ratio', y = 'TC/IS response ratio')
    p

#' ## Plot: Internal standard recovery

#+ results='asis'
cat('Mean recovery = ', signif(mean_is_recovery, 2))
#'

#+ fig.width = 7, fig.height = 4, message = FALSE

    df_spl$SampleID <- factor(df_spl$SampleID, levels = df_spl$SampleID)
    p <-    ggplot(df_spl) +
            geom_point(aes(SampleID, IS_Recovery), size = 3, color = 'dark green') +
            geom_hline(yintercept = mean_is_recovery, color = 'orange', size = 2) +
            geom_smooth(aes(1:length(df_spl$SampleID), IS_Recovery)) +
            theme_bw() +
            theme(axis.text.x = element_text(angle = 90)) +
            labs(x = 'Sample', y = 'Internal standard recovery (%)')
    p

#' ## Plot: Target compound concentration
#+ fig.width = 7, fig.height = 4

    p <-    ggplot(df_spl) +
            geom_point(aes(SampleID, Measured_TC_Conc), size = 3, color = 'red') +
            geom_hline(yintercept = mean_tc_conc, color = 'dark red', size = 2) +
            theme_bw() +
            theme(axis.text.x = element_text(angle = 90)) +
            labs(x = 'Sample', y = 'Target compound concentration')
    p

#' # Corrected calibration

# Use accuracy values to split calibration

    df_cal_1 <- df_cal %>%
                filter(Accuracy_Pct >= 110 | Accuracy_Pct <= 90) %>%
                select('SampleID', 'TC_Response', 'IS_Response', 
                    'TC_Conc', 'Response_Ratio', 'Conc_Ratio')

    df_cal_2 <- df_cal %>%
                filter(Accuracy_Pct < 110 & Accuracy_Pct > 90) %>%
                select('SampleID', 'TC_Response', 'IS_Response', 
                    'TC_Conc', 'Response_Ratio', 'Conc_Ratio')
    
    # To use TC_Conc to split calibration
    # df_cal_1 <- df_cal %>%
    #     filter(TC_Conc <= 1) %>%
    #     select('SampleID', 'TC_Response', 'IS_Response', 
    #            'TC_Conc', 'Response_Ratio', 'Conc_Ratio')
    # 
    # df_cal_2 <- df_cal %>%
    #     filter(TC_Conc > 1) %>%
    #     select('SampleID', 'TC_Response', 'IS_Response', 
    #            'TC_Conc', 'Response_Ratio', 'Conc_Ratio')

    selection_value <- last(df_cal_1$Response_Ratio) 

    df_spl_1 <- df_spl %>%
                filter(Response_Ratio <= selection_value) %>%
                select('SampleID', 'TC_Response', 'IS_Response', 'Response_Ratio')

    df_spl_2 <- df_spl %>%
                filter(Response_Ratio > selection_value) %>%
                select('SampleID', 'TC_Response', 'IS_Response', 'Response_Ratio')

# Quantificaton of first set

    # Linear model of calibration curve
    fit_1       <- lm(Response_Ratio ~ Conc_Ratio, data = df_cal_1)
    intercept_1 <- coefficients(fit_1)[1]
    slope_1     <- coefficients(fit_1)[2]

    # Calibration accuracy
    df_cal_1 <- df_cal_1 %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept_1) / slope_1) * intstd_conc) %>%
                mutate(Accuracy_Pct = abs((((Measured_TC_Conc - TC_Conc) / TC_Conc) * 100) - 100))

    r_squared_1         <- summary(fit_1)$r.squared
    mean_is_response_1  <- mean(df_cal_1$IS_Response)

    # Quantification
    df_spl_1 <- df_spl_1 %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept_1) / slope_1) * intstd_conc) %>%
                mutate(Measured_Conc_Ratio = Measured_TC_Conc / intstd_conc) %>%
                mutate(IS_Recovery = (IS_Response / mean_is_response_1) * 100)

# Quantificaton of second set

    # Linear model of calibration curve
    fit_2       <- lm(Response_Ratio ~ Conc_Ratio, data = df_cal_2)
    intercept_2 <- coefficients(fit_2)[1]
    slope_2     <- coefficients(fit_2)[2]

    # Calibration accuracy
    df_cal_2 <- df_cal_2 %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept_2) / slope_2) * intstd_conc) %>%
                mutate(Accuracy_Pct = abs((((Measured_TC_Conc - TC_Conc) / TC_Conc) * 100) - 100))

    r_squared_2         <- summary(fit_2)$r.squared
    mean_is_response_2  <- mean(df_cal_2$IS_Response)

    # Quantification
    df_spl_2 <- df_spl_2 %>%
                mutate(Measured_TC_Conc = ((Response_Ratio - intercept_2) / slope_2) * intstd_conc) %>%
                mutate(Measured_Conc_Ratio = Measured_TC_Conc / intstd_conc) %>%
                mutate(IS_Recovery = (IS_Response / mean_is_response_2) * 100)

#' ## Table: Set 1 calibration accuracy

    df_cal_1_print <-   df_cal_1 %>%
                        select(-SampleID) %>%
                        select(TC_Conc, everything())

    kable(df_cal_1_print, digits = c(1, 0, 0, 2, 2, 1, 0))

#' ## Table: Set 2 calibration accuracy

    df_cal_2_print <-   df_cal_2 %>%
                        select(-SampleID) %>%
                        select(TC_Conc, everything())

    kable(df_cal_2_print, digits = c(1, 0, 0, 2, 2, 1, 0))

#' ## Plot: Set 1 calibration

#+ results='asis'
cat('R^2^ = ', signif(r_squared_1, 4))
#'

#+ fig.width = 7, fig.height = 5

    df_cal_1_tmp <- df_cal_1 %>%
                    select(Conc_Ratio, Response_Ratio) %>%
                    mutate(Category = 'Calibration')

    df_spl_1_tmp <- df_spl_1 %>%
                    select(Measured_Conc_Ratio, Response_Ratio) %>%
                    rename(Conc_Ratio = Measured_Conc_Ratio) %>%
                    mutate(Category = 'Sample')

    df_uncorrected_1 <- bind_rows(df_spl_1_tmp, df_cal_1_tmp)

    p <-    ggplot(df_uncorrected_1) + 
            geom_point(aes(Conc_Ratio, Response_Ratio, 
                color = factor(Category), shape = factor(Category), size = factor(Category))) +
            geom_abline(intercept = intercept_1, slope = slope_1, color = 'blue') +
            theme_bw() +
            scale_color_manual(values = c('blue', 'red')) +
            scale_shape_manual(values = c(16,17)) +
            scale_size_manual(values = c(3,4)) +
            labs(x = 'TC/IS concentration ratio', y = 'TC/IS response ratio')
    p 

#' ## Plot: Set 2 calibration

#+ results='asis'
cat('R^2^ = ', signif(r_squared_2, 4))
#'

#+ fig.width = 7, fig.height = 5

    df_cal_2_tmp <- df_cal_2 %>%
                    select(Conc_Ratio, Response_Ratio) %>%
                    mutate(Category = 'Calibration')

    df_spl_2_tmp <- df_spl_2 %>%
                    select(Measured_Conc_Ratio, Response_Ratio) %>%
                    rename(Conc_Ratio = Measured_Conc_Ratio) %>%
                    mutate(Category = 'Sample')

    df_uncorrected_2 <- bind_rows(df_spl_2_tmp, df_cal_2_tmp)

    p <-    ggplot(df_uncorrected_2) + 
            geom_point(aes(Conc_Ratio, Response_Ratio, 
                color = factor(Category), shape = factor(Category), size = factor(Category))) +
            geom_abline(intercept = intercept_2, slope = slope_2, color = 'blue') +
            theme_bw() +
            scale_color_manual(values = c('blue', 'red')) +
            scale_shape_manual(values = c(16,17)) +
            scale_size_manual(values = c(3,4)) +
            labs(x = 'TC/IS concentration ratio', y = 'TC/IS response ratio')
    p 

#' # Comparison
#' Uncorrected, corrected, and MassHunter results are compared, and sorted by percent difference between the uncorrected and corrected results. 

    df_spl_1 <- mutate(df_spl_1, Set = 'Set 1')
    df_spl_2 <- mutate(df_spl_2, Set = 'Set 2')
    df_spl_c <- bind_rows(df_spl_1, df_spl_2)

    # Check for zero responses and set outputs to zero
    df_spl_c <-     df_spl_c %>%
                    mutate(Measured_TC_Conc = replace(Measured_TC_Conc, TC_Response == 0, 0)) %>%
                    mutate(Measured_Conc_Ratio = replace(Measured_Conc_Ratio, TC_Response == 0, 0))

    # Original data
    df_spl_uncorrected <-   df_spl %>%
                            select(SampleID, Measured_TC_Conc) %>%
                            rename(TC_Conc_Uncorrected = Measured_TC_Conc)
    df_spl_uncorrected$SampleID <- as.character(df_spl_uncorrected$SampleID)

    # Corrected data
    df_spl_corrected <-     df_spl_c %>%
                            select('SampleID', 'IS_Recovery', 'Measured_TC_Conc', 'Set') %>%
                            rename(TC_Conc_Corrected = Measured_TC_Conc)
    df_spl_corrected$SampleID <- as.character(df_spl_corrected$SampleID)

    # Merge original, corrected, and MassHunter data
    df_spl_comparison <-    df_spl_uncorrected %>%
                            full_join(df_spl_corrected, by = 'SampleID') %>%
                            full_join(df_spl_masshunter, by = 'SampleID') %>%
                            mutate(Difference_Pct = abs((TC_Conc_Corrected - TC_Conc_Uncorrected) / TC_Conc_Uncorrected) * 100) %>%
                            select('SampleID', 'IS_Recovery', 'TC_Conc_Uncorrected', 
                                'TC_Conc_MassHunter', 'TC_Conc_Corrected', 'Set', 'Difference_Pct') %>%
                            arrange(desc(Difference_Pct))
    
#' ## Table: Correction result
    
    kable(df_spl_comparison, digits = c(0, 0, 4, 4, 4, 1))

#' # Prepare and export results

    # Finalize units
    # Arrange in order of original sequence
    df_spl_corrected <- df_spl_corrected %>%
                        rename(Nicotine_Conc_ng_per_mL = TC_Conc_Corrected) %>%
                        mutate(Sample_Vol_mL = extract_vol) %>%
                        mutate(Nicotine_Mass_ng = Nicotine_Conc_ng_per_mL * Sample_Vol_mL) %>%
                        mutate(Batch = file_name) %>%
                        full_join(df_spl_order, by = 'SampleID') %>%
                        arrange(Order) %>%
                        select('Batch', 'SampleID', 'IS_Recovery', 'Nicotine_Conc_ng_per_mL', 
                            'Sample_Vol_mL', 'Nicotine_Mass_ng')
                        
#' ## Table: Results

    kable(df_spl_corrected, digits = c(0, 0, 0, 2, 1, 2))

# Export results 
    df_spl_corrected$IS_Recovery <- round(df_spl_corrected$IS_Recovery, 0)
    write_excel_csv(df_spl_corrected, str_c(file_name, ' Result.csv'))
