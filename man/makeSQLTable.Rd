% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/QCMemo_script.R
\name{makeSQLTable}
\alias{makeSQLTable}
\title{Generate a table from a SQL query}
\usage{
makeSQLTable(connection = con, query)
}
\arguments{
\item{connection}{A DBIConnection object, as returned by
\code{\link[DBI]{dbConnect}}.}

\item{query}{A character string containing a SQL query.}
}
\value{
A \code{\link[flextable]{flextable}} object.
}
\description{
The function \code{makeSQLTable} accepts a SQL query and displays the results
in a table.
}
\note{
The QC memo template contains a chunk of code, just after the call to
  \code{loadSASdf()}, that automatically loads all the Step 1 and Step 2
  datasets into a SQL database so that they can be accessed by queries. For
  more information, see \code{\link[DBI]{dbWriteTable}}.
}
