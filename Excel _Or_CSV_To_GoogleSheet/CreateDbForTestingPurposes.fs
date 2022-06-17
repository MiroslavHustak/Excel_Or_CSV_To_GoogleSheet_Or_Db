module CreateDbForTestingPurposes

open System
open System.Data.SqlClient //.NET

open Excel
open ROP_Functions

// REDUNDANT CODE (for testing and learning purposes only)
// Pouze pro vyzkouseni si SqlClient z .NET

let private createDbElement connString queryString = 

    use myConn = new SqlConnection (connString) // myConn.Close() ?
    let myFunction x =    
        use myCommand = new SqlCommand(queryString, myConn) 
        do myConn.Open() 
        do myCommand.ExecuteNonQuery() |> ignore
    let results =
        let ropResults = tryWith myFunction (fun x -> myConn.Close()) (fun ex -> failwith)
        ropResults |> deconstructor () String.Empty
    results  

let deleteDataAndFillInTable() = 

    let myConnString = @"Data Source=Misa\SQLEXPRESS;Initial Catalog=DGSada;Integrated Security=True"       

    let deleteDataInTable() = 
        let myQueryString = "DELETE FROM SOAL"                         
        createDbElement <| myConnString <| myQueryString

    let fillInTable() = 
        let myQueryString = "
                            INSERT INTO SOAL (Id,
                                             [Pracovní značení],
                                             [Digitalizační sada],
                                             [Archiv],
                                             [Fond],
                                             [Číslo NAD],
                                             [Číslo pomůcky],
                                             [Inventární číslo],
                                             [Signatura],
                                             [Číslo kartonu],
                                             [Upřesňující indentifikátor],
                                             [Regest],
                                             [Datace vzniku],
                                             [Poznámka])
                            VALUES (@val01, @val02, @val03, @val04, @val05, @val06, @val07,
                                    @val08, @val09, @val10, @val11, @val12, @val13, @val14);   
                            "    
        
        use myConn = new SqlConnection (myConnString) // myConn.Close() ?
       
        let myFunction x =    
            let (header, rows) = readDataFromExcel()            
            do myConn.Open()
            rows |> Array.iteri(fun i row -> 
                                            use myCommand = new SqlCommand(myQueryString, myConn)
                                            myCommand.Parameters.AddWithValue("@val01", i + 1)    |> ignore
                                            myCommand.Parameters.AddWithValue("@val02", row.[0])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val03", row.[1])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val04", row.[2])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val05", row.[3])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val06", row.[4])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val07", row.[5])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val08", row.[6])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val09", row.[7])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val10", row.[8])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val11", row.[9])  |> ignore
                                            myCommand.Parameters.AddWithValue("@val12", row.[10]) |> ignore
                                            myCommand.Parameters.AddWithValue("@val13", row.[11]) |> ignore
                                            myCommand.Parameters.AddWithValue("@val14", row.[12]) |> ignore 
                                            myCommand.ExecuteNonQuery() |> ignore
                               ) 
        let results = 
            let ropResults = tryWith myFunction (fun x -> myConn.Close()) (fun ex -> failwith)
            ropResults |> deconstructor () String.Empty
        results 
    
    do deleteDataInTable()
    do fillInTable()

