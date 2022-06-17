open System

open Csv
open Excel
open CreateDb
open StartForTestingPurposes
open CreateDbForTestingPurposes

[<EntryPoint>]
let main argv = 

    do start1() |> ignore  //zrusit Console.ReadLine() v readLine module excel //pouze pro testovani async a multithreading
   
    //do printfn "Tvorba nové db (zadat kód)" 
    //do printfn "Tvorba nové db tabulky (zadat kód)"
    //do printfn "Excel nebo csv -> Google Sheets a db (zadat e nebo c)"
   
    let leaveThisApp() = 
        do printfn "Tož to by bylo vše ohledně přenosu dat z Excelu nebo csv souboru do Google Sheet a db ..."
        do printfn "Press any key to end this application"
        do Console.ReadKey() |> ignore 

    let crossroads() = 
        match Console.ReadLine() with
        | "74764" -> do createDb()                 |> ignore 
                     do leaveThisApp() 
        | "74283" -> do createTable()              |> ignore 
                     do leaveThisApp()  
        | "74766" -> do deleteDataAndFillInTable() |> ignore 
                     do leaveThisApp()  
        | "e"     -> do excel()                    |> ignore 
                     do leaveThisApp()  
        | "c"     -> do csv()                      |> ignore
                     do leaveThisApp() 
        | _       -> do leaveThisApp() 
    
    crossroads()
    
    0 // return an integer exit code


    