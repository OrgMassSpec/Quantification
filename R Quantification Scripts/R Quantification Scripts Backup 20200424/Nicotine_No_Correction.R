#' ---
#' title: "Project: | Matrix: | Analyte: Nicotine"
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
    options("width"=200)
    options(scipen=999)

# source('Nicotine_No_Correction.R')
# rmarkdown::render('Nicotine_No_Correction.R')

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

    # instrumental sequence name
    seq_name    = '???'

    # input (.xlsx) and generated output (.csv) file name
    file_name   = 'Data - Nicotine No Correction'    

    intstd_conc = 5     # internal standard concentration (ng/mL)
    extract_vol = 3     # acetonitrile extraction volume (mL)

#+ results='asis'
cat('__Sequence Name:__ ', seq_name, '.', sep = '')
#'

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
    x <- read_excel(str_c(file_name, '.xlsx'), sheet = 'Sheet1')

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

#' # Comparison
#' Uncorrected and MassHunter results are compared. 

    # Uncorrected data
    df_spl_uncorrected <-   df_spl %>%
                            select(SampleID, IS_Recovery, Measured_TC_Conc) %>%
                            rename(TC_Conc_Uncorrected = Measured_TC_Conc)
    df_spl_uncorrected$SampleID <- as.character(df_spl_uncorrected$SampleID)

    # Merge original, corrected, and MassHunter data
    df_spl_comparison <-    df_spl_uncorrected %>%
                            full_join(df_spl_masshunter, by = 'SampleID') %>%
                            select('SampleID', 'IS_Recovery', 'TC_Conc_Uncorrected', 
                                'TC_Conc_MassHunter')
    
#' ## Table: Correction result
    
    kable(df_spl_comparison, digits = c(0, 0, 4, 4))

#' # Prepare and export results

    # Finalize units
    # Arrange in order of original sequence
    df_spl <-   df_spl %>%
                rename(Nicotine_Conc_ng_per_mL = Measured_TC_Conc) %>%
                mutate(Sample_Vol_mL = extract_vol) %>%
                mutate(Nicotine_Mass_ng = Nicotine_Conc_ng_per_mL * Sample_Vol_mL) %>%
                mutate(Batch = file_name) %>%
                full_join(df_spl_order, by = 'SampleID') %>%
                arrange(Order) %>%
                select('Batch', 'SampleID', 'IS_Recovery', 'Nicotine_Conc_ng_per_mL', 
                    'Sample_Vol_mL', 'Nicotine_Mass_ng')
                        
#' ## Table: Results

    kable(df_spl, digits = c(0, 0, 0, 2, 1, 2))

# Export results 
    df_spl$IS_Recovery <- round(df_spl$IS_Recovery, 0)
    write_excel_csv(df_spl, str_c(file_name, ' Result.csv'))
