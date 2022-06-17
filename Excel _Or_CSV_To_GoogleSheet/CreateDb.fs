module CreateDb

open System
open System.Data.SqlClient //.NET

open ROP_Functions

let private createDbElement connString queryString = 
    use myConn = new SqlConnection (connString) 
    let myFunction x =    
        use myCommand = new SqlCommand(queryString, myConn) 
        myConn.Open() 
        myCommand.ExecuteNonQuery() |> ignore   
    let results = 
        let ropResults = tryWith myFunction (fun x -> myConn.Close()) (fun ex -> failwith)
        ropResults |> deconstructor () String.Empty
    results  

let createDb() =  
    let myConnString = @"Server=Misa\SQLEXPRESS;Integrated security=SSPI;database=master" 
    let myQueryString = "CREATE DATABASE DGSada"
    createDbElement <| myConnString <| myQueryString

let createTable() = 
    let myConnString = @"Data Source=Misa\SQLEXPRESS;Initial Catalog=DGSada;Integrated Security=True"
    let myQueryString = "
                        CREATE TABLE SOAL 
                        (
                            Id int NOT NULL PRIMARY KEY,
                            [Pracovní značení] nvarchar(255),
                            [Digitalizační sada] nvarchar(255),
                            [Archiv] nvarchar(255),
                            [Fond] nvarchar(255),
                            [Číslo NAD] nvarchar(255),
                            [Číslo pomůcky] nvarchar(255),
                            [Inventární číslo] nvarchar(255),
                            [Signatura] nvarchar(255),
                            [Číslo kartonu] nvarchar(255),
                            [Upřesňující indentifikátor] nvarchar(255),
                            [Regest] nvarchar(255),
                            [Datace vzniku] nvarchar(255),
                            [Poznámka] nvarchar(255)
                        );   
                        "    
    createDbElement <| myConnString <| myQueryString
