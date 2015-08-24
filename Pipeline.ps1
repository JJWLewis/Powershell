#Pipeline stuff

A pipeline is a series of commands connected by pipeline operators
   (|)(ASCII 124). Each pipeline operator sends the results of the preceding
    command to the next command.
     
    You can use pipelines to send the objects that are output by one command
    to be used as input to another command for processing. And you can send the
    output of that command to yet another command. The result is a very powerful
    command chain or "pipeline" that is comprised of a series of simple commands. 


#convert to HTML, then save
get-process | ConvertTo-HTML -Property name,status | out-file C:\Users\jay.lewis\Desktop\Powershell\test.html

#if you don't know what something does
get-process | ConvertTo-HTML -Property name,status | out-file C:\Users\jay.lewis\Desktop\Powershell\test.html -WhatIf


