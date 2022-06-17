module MyTypes

open System
open FSharp.Interop.Excel

open FSharp.Data           //F# SqlCommandProvider, SqlProgrammabilityProvider
open FSharp.Data.Sql       //F# SqlDataProvider
open FSharp.Data.SqlClient //F# ResultType.DataReader

let [<Literal>] TypeProviderConn = @"Data Source=Misa\SQLEXPRESS;Initial Catalog=DGSada;Integrated Security=True" //Values that are intended to be constants can be marked with the Literal attribute. //quli type provider uz pri compile time

type MyCSV = CsvProvider< @"e:\E\Mirek po osme hodine a o vikendech\Pruvodky Litomerice\CSV Digitalizacni sada\LT-00001 az LT-00204.csv", ";", 0> //0 znamena, ze uplne vsechny radky budou mit dedukovany typ (default je jen prvnich 1000 radku)
//type MyCSV = CsvProvider< @"e:\Mirek po osme hodine a o vikendech\LT-00001 az LT-04640.csv", ";", 0> 

//vyzaduje mimo jine package ExcelDataReader.DataSet
type ExcelProvider = ExcelFile<"e:\E\Mirek po osme hodine a o vikendech\Pruvodky Litomerice\XLSXnew Digitalizacni sada\LT-00001 az LT-00204 DGSada 01-06-2021.xlsx", SheetName="Table1", HasHeaders = false, ForceString = true>

type MySOALTable = SqlProgrammabilityProvider<TypeProviderConn>

type ContextClass = SqlDataProvider<ConnectionString = TypeProviderConn, UseOptionTypes = true> // pro SqlDataProvider //ORM

type Result<'TSuccess,'TFailure> =
| Success of 'TSuccess
| Failure of 'TFailure

type TaskResults =   
| ExcelInt of excel: unit  //abych mohl pouzit jen Async.Start, musi byt typ unit (jinak musim pouzit Async.Parallel, Async.RunSynchronously)
| CsvInt of csv: unit
    