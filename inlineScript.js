
    function WriteToFile(passForm) {
   
       set fso = CreateObject("Scripting.FileSystemObject");  
       set s = fso.CreateTextFile("test.csv", True);
       s.writeline("TEST THIS");
       s.Close();
    }


