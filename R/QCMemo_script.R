# LOAD REQUIRED LIBRARIES (INSTALL IF NOT ALREADY)

#' @import usethis

usethis::use_package('usethis')
usethis::use_package('tidyr')
usethis::use_package('dplyr')
usethis::use_package('flextable')
usethis::use_package('haven')
usethis::use_package('DBI')
usethis::use_package('RSQLite')
usethis::use_package('odbc')
usethis::use_package('readxl')
usethis::use_package('stats')

#' @import tidyr
#' @import dplyr
#' @import flextable
#' @importFrom haven read_sas
#' @import DBI
#' @importFrom RSQLite SQLite
#' @import odbc
#' @importFrom readxl excel_sheets read_xlsx
#' @importFrom stats ftable reshape

con <- dbConnect(RSQLite::SQLite(), dbname = ':memory:')

# LOAD CUSTOM FUNCTIONS

#' Generate the QC log
#'
#' This function organizes the information from the doc block at the beginning
#' of the template into a table detailing all the programs, datasets, and other
#' files covered in the QC memo.
#'
#' @param rootdir The root directory for all programs and data. This directory
#'   should contain subfolders for programs and SAS datasets (e.g. "prg" and
#'   "sd7" folders).
#'
#' @param rawprg The name of the raw step program.
#' @param rawdat A vector containing the names of all the raw datasets.
#'
#' @param s1prg The name of the Step 1 program.
#' @param s1dat A vector containing the names of all the Step 1 datasets. \emph{Must
#'   end in .sas7bdat and be uncompressed.}
#'
#' @param s2prg The name of the Step 2 program.
#' @param s2dat The name of the Step 2 dataset. \emph{Must end in .sas7bdat and be uncompressed.}
#' @param s2mac The name of the Step 2 macro Excel worksheet. \emph{Must end in .xlsm.}
#'
#' @param prgmer_raw The raw step programmer.
#' @param prgmer_s1 The Step 1 programmer.
#' @param prgmer_s2 The Step 2 programmer.
#' @param reviewers The names of the QC memo reviewers
#' @param QC_date The date of the QC memo review.
#'
#' @return A \code{\link[flextable]{flextable}} object.
#' @export
makeQCLog <- function(rootdir, rawprg, rawdat, s1prg, s1dat, s2prg, s2dat, s2mac,
                      prgmer_raw, prgmer_s1, prgmer_s2, reviewers, QC_date) {

  nrawdat <- length(rawdat)
  ns1dat <- length(s1dat)

  if (!is.null(s2dat)) {
    ns2dat <- length(s2dat)
    total_rows <- nrawdat + ns1dat + ns2dat + 4
  }
  else {
    total_rows <- nrawdat + ns1dat + 2
  }

  if (!is.null(s2dat)) {
    tabl2 <- data.frame(Stage = c(rep('Raw Step', 1 + nrawdat),
                                  rep('Step 1', 1 + ns1dat),
                                  rep('Step 2', 1 + ns2dat + 1)),
                        Type  = c('Program', rep('Dataset', nrawdat),
                                  'Program', rep('Dataset', ns1dat),
                                  'Program', rep('Dataset', ns2dat), 'Macro'),
                        File  = c(rawprg, rawdat,
                                  s1prg, s1dat,
                                  s2prg, s2dat, s2mac),
                        Programmer = c(rep(prgmer_raw, 1 + nrawdat),
                                       rep(prgmer_s1, 1 + ns1dat),
                                       rep(prgmer_s2, 1 + ns2dat + 1)),
                        Reviewers = rep(reviewers, total_rows),
                        Date = rep(QC_date, total_rows))
  }

  qclog_ft <- flextable(data = tabl2) %>%
    merge_v(j = c('Stage', 'Type', 'File', 'Programmer', 'Reviewers', 'Date')) %>%
    theme_box() %>%
    fontsize(size = 9)
  qclog_ft <- width(qclog_ft, width = dim_pretty(qclog_ft)$widths)
  qclog_ft
}

fixFilepath <- function(path) {
  envvars <- c('!PSE', '!K12', '!LWWC', '!YCE', '!FNC', '!CDI', '!XP', '!ARC')
  folders <- c('pse', 'k12', 'lwwc', 'yce', 'fnc', 'cdi', 'xproject', 'archives')

  for (i in 1:length(envvars)) {
    if (length(grep(pattern = envvars[i], path, ignore.case = TRUE))) {
      path <- gsub(pattern = envvars[i], replacement = folders[i], path, ignore.case = TRUE)
    }
  }

  return(path)
}

#' Load SAS datasets
#'
#' The \code{loadSASdf} function reads in the Step 1 and Step 2 files specified
#' in the doc block and automatically assigns them names (unless names are
#' user-specified).
#'
#' @inheritParams makeQCLog
#' @param dirabb The name of the directory (within the root directory) where the SAS datasets are stored.
#' @param suffix A suffix to add to the end of all the filenames (before the
#'   file extension). For example, if the doc block lists the uncompressed
#'   datasets and you would like to add "_uncomp" to all the filenames, setting
#'   \code{suffix = "_uncomp"} would change "MyData.sas7bdat" to
#'   "MyData_uncomp.sas7bdat".
#' @param s1names A vector of names for the Step 1 datasets. \emph{Must be the
#'   same length as \code{s1dat}}
#' @param s2names A vector of names for the Step 2 datasets. \emph{Must be the
#'   same length as \code{s2dat}}
#'
#' @return A named \code{list} containing all the Step 1 and Step 2 datasets.
#' @export
loadSASdf <- function(rootdir, dirabb = 'SD7', suffix = '', s1dat, s1names = NULL, s2dat, s2names = NULL) {

  rootdir <- fixFilepath(rootdir)

  ns1dat <- length(s1dat)

  if (!is.null(s2dat)) {
    ns2dat <- length(s2dat)
  }

  if (is.null(s1names)) {
    s1names <- paste0('s1_', seq(1, length(s1dat), 1))
  }

  if (is.null(s2names)) {
    s2names <- 's2'
  }

  dirabb <- paste0('/', dirabb, '/')
  suffix <- paste0(suffix, '.')

  dflist <- list()

  for (i in 1:length(s1dat)) {
    assign(s1names[i], read_sas(paste0(rootdir, dirabb, gsub('.', suffix, s1dat[i], fixed = TRUE))))
    dflist[[i]] <- eval(parse(text = s1names[i]))
  }

  if (!is.null(s2dat)) {
    for (j in 1:length(s2dat)) {
      assign(s2names[j], read_sas(paste0(rootdir, dirabb, gsub('.', suffix, s2dat[j], fixed = TRUE))))
      dflist[[i+j]] <- eval(parse(text = s2names[j]))
    }
  }

  if (!is.null(s2dat)) {
    names(dflist) <- c(s1names, s2names)
  }
  else {
    names(dflist) <- c(s1names)
  }

  return(dflist)
}

#' Create sample size table
#'
#' The function \code{makeSSTable} creates the standard sample size table
#' included in all QC memos (i.e., cross-tabulated by cohort, context, and
#' research group).
#'
#' @param inputds A dataset with one row per unit (i.e. a Step 2 dataset).
#' @param cohort_var The name of the cohort variable on \code{inputds}.
#' @param context_var The name of the context variable on \code{inputds}.
#' @param racode_var The name of the research group variable on \code{inputds}.
#'
#' @return A \code{\link[flextable]{flextable}} object.
#' @export
makeSSTable <- function(inputds, cohort_var = 'cohort', context_var = 'context', racode_var = 'racode') {
  sstab <- ftable(inputds[[cohort_var]],
                  inputds[[context_var]],
                  inputds[[racode_var]],
                  row.vars = c(1:2),
                  col.vars = 3)

  ssdf <- data.frame(sstab)
  colnames(ssdf) <- c(cohort_var, context_var, racode_var, 'Freq')
  ssdf <- ssdf[order(ssdf[[cohort_var]], ssdf[[context_var]]),]
  ssdf[,cohort_var] <- as.character(ssdf[,cohort_var])
  ssdf[,context_var] <- gsub('_', ' ', as.character(ssdf[,context_var]))
  ssdf <- reshape(data = ssdf,
                  timevar = racode_var, idvar = c(cohort_var, context_var),
                  times = levels(ssdf[,racode_var]), direction = 'wide')
  ssdf <- rbind(ssdf, c('Total', '', apply(ssdf[,c(3:ncol(ssdf))], 2, sum)))

  for (i in 3:ncol(ssdf)) {
    ssdf[,i] <- as.numeric(ssdf[,i])
  }

  ssdf$Total <- apply(X = ssdf[,3:ncol(ssdf)], MARGIN = 1, FUN = sum)

  colnames(ssdf) <- c('Cohort', 'Context', gsub('Freq.', '', colnames(ssdf)[3:(ncol(ssdf)-1)], fixed = TRUE), 'Total')

  ss_ft <- flextable(data = ssdf) %>%
    merge_v(j = 1) %>%
    theme_vanilla()
  ss_ft <- width(ss_ft, width = dim_pretty(ss_ft)$widths)
  ss_ft
}

#' Generate list of Step 2 outcomes
#'
#' The function \code{step2Contents} scrapes the Step 2 Excel worksheet to print
#' a table showing the stem and label for all Step 2 outcomes.
#'
#' @param dirabb The name of the directory (within the root directory) where the
#'   Step 2 Excel worksheet is stored.
#' @inheritParams makeQCLog
#'
#' @return A \code{\link[flextable]{flextable}} object.
#' @export
step2Contents <- function(rootdir, dirabb = 'PRG', s2mac) {
  rootdir <- fixFilepath(rootdir)
  dir <- paste0(rootdir, '/', dirabb, '/', s2mac)
  sheets <- excel_sheets(dir)
  sheets <- sheets[!(sheets %in% c('Terms', 'Validate'))]

  contents <- data.frame(Stem = character(),
                         Label = character())

  for (s in sheets) {
    assign(s, read_xlsx(path = dir, sheet = s, col_names = TRUE)[,1:2])
    contents <- rbind(contents, eval(parse(text = s)))
  }

  cont_ft <- flextable(contents) %>%
    theme_vanilla() %>%
    align(align = 'left', part = 'all') %>%
    fontsize(size = 9)
  cont_ft <- width(cont_ft, width = dim_pretty(cont_ft)$widths)
  cont_ft

  return(cont_ft)
}

#' Generate a frequency table
#'
#' The function \code{freqTable} is somewhat analagous to \code{PROC FREQ} in
#' SAS. The functionality is slightly different and more limited, but the
#' results are aesthetically similar.
#'
#' @param variables A vector of names indicating the variables you would like to
#'   be cross-tabulated.
#' @param varnames A named \code{list} giving key-value pairs for how you would
#'   like the \code{variables} to be labeled on the table.
#' @param n Logical. Should frequencies be displayed?
#' @param percent Logical. Should percents be displayed?
#' @param perc_digits A number indicating how many decimal places should be
#'   displayed for the percentages.
#' @param bg The background color for the columns containing the calculated
#'   statistics. Can be given as either a name (e.g. \code{"dodgerblue4"}) or
#'   using the \code{\link[grDevices]{rgb}} function.
#' @inheritParams makeSSTable
#'
#' @return A \code{\link[flextable]{flextable}} object.
#'
#' @examples
#' # Simple cross-tabulation using default options
#' freqTable(inputds    = mtcars,
#'           variables  = c('cyl', 'gear', 'carb'))
#'
#' # Using options to customize table further
#' freqTable(inputds    = mtcars,
#'           variables  = c('cyl', 'gear', 'carb'),
#'           varnames   = list(cyl = 'Cylinders', gear = 'Forward gears', carb = 'Carburetors'),
#'           n          = FALSE,
#'           perc_digits = 0,
#'           bg         = rgb(.2, .1, .2, .3))
#' @export
freqTable <- function(inputds, variables, varnames = variables, n = TRUE, percent = TRUE, perc_digits = 1, bg = 'aliceblue') {
  df <- inputds %>%
    group_by_at(vars(all_of(variables))) %>%
    summarise(N = n(), Percent = 100*(N/nrow(inputds))) %>%
    arrange_at(vars(all_of(rev(variables))))

  if (n & percent) {
    tbl <- flextable(df) %>%
      theme_vanilla() %>%
      colformat_num(j = 'Percent', digits = perc_digits, suffix = '%') %>%
      bg(j = c('N', 'Percent'), bg = bg, part = 'all') %>% set_header_labels(values = varnames)
    tbl <- width(tbl, width = dim_pretty(tbl)$widths)
  }

  if (n & !percent) {
    tbl <- flextable(df[,1:(ncol(df)-1)]) %>%
      theme_vanilla() %>%
      bg(j = 'N', bg = bg, part = 'all') %>% set_header_labels(values = varnames)
    tbl <- width(tbl, width = dim_pretty(tbl)$widths)
  }

  if (!n & percent) {
    tbl <- flextable(df[,c(1:length(variables), ncol(df))]) %>%
      theme_vanilla() %>%
      colformat_num(j = 'Percent', digits = perc_digits, suffix = '%') %>%
      bg(j ='Percent', bg = bg, part = 'all') %>% set_header_labels(values = varnames)
    tbl <- width(tbl, width = dim_pretty(tbl)$widths)
  }

  if (!n & !percent) {
    tbl <- flextable(df[,1:length(variables)]) %>%
      theme_vanilla() %>% set_header_labels(values = varnames)
    tbl <- width(tbl, width = dim_pretty(tbl)$widths)
  }
  tbl
}

#' Generate a table from a SQL query
#'
#' The function \code{makeSQLTable} accepts a SQL query and displays the results
#' in a table.
#'
#' @param connection A DBIConnection object, as returned by
#'   \code{\link[DBI]{dbConnect}}.
#' @param query A character string containing a SQL query.
#'
#' @return A \code{\link[flextable]{flextable}} object.
#'
#' @note The QC memo template contains a chunk of code, just after the call to
#'   \code{loadSASdf()}, that automatically loads all the Step 1 and Step 2
#'   datasets into a SQL database so that they can be accessed by queries. For
#'   more information, see \code{\link[DBI]{dbWriteTable}}.
#' @export
makeSQLTable <- function(connection = con, query) {
  results <- dbGetQuery(connection, query)
  outtbl <- flextable(results) %>%
    theme_vanilla()
  outtbl <- width(outtbl, width = dim_pretty(outtbl)$widths)
  outtbl
}

#' Generate a summary table
#'
#' The function \code{summaryTable} is roughly analogous to \code{PROC MEANS} in
#' SAS. The functionality is slightly more limited than \code{PROC MEANS}, but
#' the results are aesthetically similar.
#'
#' @param inputds A dataset to be summarized.
#' @param class A vector of names indicating variables to be used for
#'   classification.
#' @param var The name of the analysis variable.
#' @param varnames A named \code{list} giving key-value pairs for how you would
#'   like the variables to be labeled on the table.
#' @param nmiss Logical. Should number of missing values be displayed?
#' @param mean Logical. Should means be displayed?
#' @param min Logical. Should minima be displayed?
#' @param max Logical. Should maxima be displayed?
#' @param digits A number indicating how many decimal places should be displayed
#'   for the mean, minimum, and maximum
#' @param date Logical. Is the analysis variable a date?
#' @inheritParams freqTable
#'
#' @return A \code{\link[flextable]{flextable}} object.
#'
#' @examples
#' # Simple table with one class variable
#' summaryTable(inputds     = mtcars,
#'              class       = 'cyl',
#'              var         = 'mpg',
#'              varnames    = list(cyl = 'Cylinders', mpg = 'Miles per gallon'))
#'
#' # Two class variables and customization through additional parameters
#' summaryTable(inputds     = mtcars,
#'              class       = c('cyl', 'gear'),
#'              var         = 'mpg',
#'              varnames    = list(cyl  = 'Cylinders',
#'                                 gear = 'Number of forward gears',
#'                                 mpg  = 'Miles per gallon'),
#'              nmiss       = FALSE,
#'              perc_digits = 0,
#'              digits      = 2,
#'              bg          = 'olivedrab1')
#' @export
summaryTable <- function(inputds, class, var, varnames = c(class, var), n = TRUE, nmiss = TRUE, percent = TRUE,
                         mean = TRUE, min = TRUE, max = TRUE, perc_digits = 1, digits = 1, bg = 'aliceblue', date = FALSE) {
  df <- inputds %>%
    group_by_at(vars(all_of(class))) %>%
    summarise(N       = n(),
              NMiss   = sum(is.na(!!as.name(var))),
              Percent = 100*(N/nrow(inputds)),
              Mean    = mean(!!as.name(var), na.rm = TRUE),
              Min     = min(!!as.name(var), na.rm = TRUE),
              Max     = max(!!as.name(var), na.rm = TRUE))

  keepcols <- c(rep(TRUE, length(class)), n, nmiss, percent, mean, min, max)
  outtbl <- flextable(df[,keepcols]) %>%
    merge_v(class) %>%
    set_header_labels(values = varnames) %>%
    theme_vanilla() %>%
    bg(j = (length(class)+1):sum(keepcols), bg = bg, part = 'all')

  outtbl <- width(outtbl, width = dim_pretty(outtbl)$widths)
  if (percent) {
    outtbl <- outtbl %>% colformat_num(j = 'Percent', digits = perc_digits, suffix = '%')
  }
  if (mean & !date) {
    outtbl <- outtbl %>% colformat_num(j = 'Mean', digits = digits)
  }
  if (min & !date) {
    outtbl <- outtbl %>% colformat_num(j = 'Min', digits = digits)
  }
  if (max & !date) {
    outtbl <- outtbl %>% colformat_num(j = 'Max', digits = digits)
  }
  outtbl
}

