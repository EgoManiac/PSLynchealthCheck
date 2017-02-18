#region Script Information 
Clear-Host 

Write-Host "----------------------------------------------------" -BackgroundColor DarkGreen 
Write-Host "|   _                       ____   ___  _ _____ " -ForegroundColor Green 
Write-Host "|  | |   _   _ _ __   ___  |___ \ / _ \/ |___ / " -ForegroundColor Green 
Write-Host "|  | |  | | | | '_ \ / __|   __) | | | | | |_ \ " -ForegroundColor Green 
Write-Host "|  | |__| |_| | | | | (__   / __/| |_| | |___) |" -ForegroundColor Green 
Write-Host "|  |_____\__, |_| |_|\___| |_____|\___/|_|____/ " -ForegroundColor Green 
Write-Host "|        |___/   " -ForegroundColor Green 
Write-Host "|" -ForegroundColor Green 
Write-Host "|"-ForegroundColor Green 
Write-Host "| System Health Check" -ForegroundColor Green 
Write-Host "| Version: 0.30" -ForegroundColor Green 
Write-Host "| Feb 2017" -ForegroundColor Green 
Write-Host "|"-ForegroundColor Green 
Write-Host "| Authors:" -ForegroundColor Green 
Write-Host "|  Chris Burns      | chris.burns@hpe.com" -ForegroundColor Green 
Write-host "|"-ForegroundColor Green 
Write-Host "----------------------------------------------------" -BackgroundColor DarkGreen 
Write-Host 
#endregion 

 


#region Variables

$colorscheme = "#00b388"
$brand2color = "#425563"
$overrideAutoStrip = $false
$charstokeepforserverid =2




$Header = @"
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<style>
.entirepage {border-width: 15px; border-style: solid; border-color: $colorscheme; vertical-align :top; padding: 15px; font-family: "verdana";}
.headerimage {width: 250px}

TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; table-layout:fixed;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; font-size: 10pt; color: white; background-color: $brand2color;}
TD {border-width: 0px; border-style: solid; border-color: black; font-size: 8pt; vertical-align :top; padding: 10px}
TR:nth-child(odd) {background-color :lightgray}
.pools td:nth-child(1) { color: $colorscheme; }
.pools td:nth-child(2) { color: $brand2color; }

.computer_pass {border-width: 2px; border-style: solid; border-color: black; vertical-align :top; padding: 10px; color: white; background-color: green; }
.computer_fail {border-width: 2px; border-style: solid; border-color: black; vertical-align :top; padding: 10px; color: white; background-color: red; }

.replica_pass {border-width: 2px; border-style: solid; border-color: black; vertical-align :top; padding: 10px; color: white; background-color: green; }
.replica_fail {border-width: 2px; border-style: solid; border-color: black; vertical-align :top; padding: 10px; color: white; background-color: red; }


.FEandMed_container { display: flex; }
.EdgeandSQL_Container { display: flex; }
.Replication_Container  { display: flex; }

.servers {min-height: 300px; border-width: 2px; border-style: solid; border-color: $brand2color; vertical-align :top; margin: 10px; width: 500px; display: inline-block; position:relative;border-top-right-radius: 7px; border-top-left-radius: 7px; border-bottom-right-radius: 2px; border-bottom-left-radius: 2px; background-color: #F0F0F0 }
.servers h2 {background:$brand2color; color:white;padding:10px; margin:0;font-size: 12pt;}
.servers h3 {padding:0px; margin:0px;font-size: 10pt; margin-top: 5px}
.serverdivcontent {padding:10px;}

.repservices {min-height: 300px; border-width: 2px; border-style: solid; border-color: $brand2color; vertical-align :top; margin: 10px; width: 720px; display: inline-block; position:relative;border-top-right-radius: 7px; border-top-left-radius: 7px; border-bottom-right-radius: 2px; border-bottom-left-radius: 2px; background-color: #F0F0F0 }
.repservices h2 {background:$brand2color; color:white;padding:10px; margin:0;font-size: 12pt;}
.repservices h3 {padding:0px; margin:0px;font-size: 10pt; margin-top: 5px}

.sqlservices {min-height: 300px; border-width: 2px; border-style: solid; border-color: $brand2color; vertical-align :top; margin: 10px; width: 280px; display: inline-block; position:relative;border-top-right-radius: 7px; border-top-left-radius: 7px; border-bottom-right-radius: 2px; border-bottom-left-radius: 2px; background-color: #F0F0F0 }
.sqlservices h2 {background:$brand2color; color:white;padding:10px; margin:0;font-size: 12pt;}
.sqlservices h3 {padding:0px; margin:0px;font-size: 10pt; margin-top: 5px}

.servergraphic { position:absolute; bottom:-10px; right:-20px; height: 50px} 

.subtitle { font-size: 8pt; font-style: oblique; }

div#poolinfo {display:none;}

div#poolinfo:target {display:block;}
</style>
"@


$CompanyLogoBase64 = "data:image/png;base64,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"
$EdgeGraphicBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGcAAABVCAIAAADbkW0IAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAXrSURBVHhe7ZxfaFtVHMd989HHvfki+LAHwRffxCd9ExxjeXAdLC9OMCCLUh9swMWVZXN9SAdrUJhgGI3QDtYipKNRyqgkZbbT1HVhaSltZo0yM0epa8qIX3pOj7fnnpt7c+45udfkfPg9ZPece869n3v+prl7rmXoHGNNBmNNBmNNBt5arVZLpS4cO/aOjojHzxYKBVpT16nWt4vVv9L5amLi3rtX7xwdnJ1YeEjTOuSQtXq9HolEuFtVHvl8ntank6X1x9OLW3D0/rW7x9Oll+K37KHGmr5WZg08mGazSatUwd87e2hE1+c3v/juARrRG+dvc3acQo21gYGT3B1qikqlQqtUAZRxOjyGGmvcvemLcrlMq1RBj1jDqJ9IDLWJtbU1WqUK/t/WMpkxtY3II36s3X345OXhuSNDBXygxXlAjbVo9PTi4iItpev4sXapsPr8R3nEi5/98PuTp7REN1ysYb7jOpcwNjY2aBGtFuZHtDjlYFVEK7Dhx9pc9RGxhnjz6sI/e89ooW1xsXbmzHs0zQPwlclkuBIURjQaFU6+QmvRL3/C8fZCYa2xs/f6aJGJ+/jmfbS4E18vfT7zYGblTyeJKq1duTLKna480Paxe6H1HSBUMzi+TFK549aANbQvpozEt0u/IcjnVy7eLq03SDlWXKzh8dIe4sDKygo5Fz2IO1dT4NmQGhlqrb3wyS3MDGh07AgaHSmK4Xc2SCbPkXOLxSKXpCmwrCE1MnxawzxwKvvzNws1jHEk7te3ceLbX90h1uwThV9r7MljW84laQr7oOHHGoYwpzUHhjz4IuJiE7/So/v0uzWoIdmEoAESa6+N/EgP7dPv1kgeJ9AxiTWEdT411tqBlsisWVulXmup1IVcbtxPcAUiummNLUEwwNFD++i1hoMkVRquQIRaa/BCstlB48IWlVjDJEuP7tP71tIzqyxe/fR7ax5Yw8qDiMNqlq08EJgH2ASK4ObZ3u+hVriveYk1LGuh7Ga5zhxxgd0VPf+Anp0NsA+9V+MXYkJr8IKeiOnSuh8gAaGXCqv0ZAs9a80ex9Olo4Oz1iOwhs6IMevIUAF7eKwtrBssfHb67qiPrNkD1mgRB0ATmwHQ0MjWyo6xxoMxDr6IuNG5dXr0MMaaAMwMmEDRbU1bE4STNYAxrs33usaaDL1gbfPRTjpf/TD7i9MvE5yir61xQCJa39jsGvkVDLcfsIax1o6ne8/g8fr8JpokPL6VmjfWJFlaf/yH5z+AcvSvNT8YazIYazIYazIYazIYazIYazIYazIYazIYazIYazIYazKE11okciLAX0i3J6TWBgZOqn0nQS1htBZyZSB01jBsWX+vjPK5DKiIptmo1+s43Zo5Eono6ObhsoZ7Zu8VNJvNZDLJZWijjNAdcS7WvE9Y/q3F42cbDfozdTllBLs47+d6JCzWoGx7m/7JFh/wTy7D1NQUSfWCbnEu1lzfN2BjkB9ricQQU4bmZleWy42TVO8IxUmUIyT4cS2ZPMdeTVZ7q8LSstksTfZBwNZGRi4zZWi29pvE6EZbtRS4KswGXJnsmqUJ0pr16rFA69r74whUzZ6WBIFZy2TGyImgUql0UxkJtGJpccFY45TZO1F3AuLoRXSIizXcj/XtWXvcuDFJzhVai8ViXH4Sw8Pnt7boT9aFS7PuBK6ZXEOnuFjD8EzT3BBa83hZgYiTVgZCYQ04iXNqrd5DOGL6UQZcrPlc5WJxRPM5gEFtd3eXlCDcEuCxYdlFMkjg81k6EcxsYA2rF7XiNCkDwVtDuIrDEbbl8og+ZSAU1hBqxQkvZnKSTvf+CYs1RCz2AfPiR5zwSlTt2wkhsoawepET1wVlQK81iXf28vl8s/nfrMqlIorFIkm1I8yPC6PJ6tBrTccVh4FD1rin5CVKJfrk+9eaH4TW/LxVi666vFzO2ZienqZVBodeazoCaxRaZXAYazIYazIYazIYazIYazIYazIos9ZoNOg3jZpR+/86y6HMWl9hrMlgrHVOq/UvlBgjXDoC1GsAAAAASUVORK5CYII="
$FrondEndGraphicBase64 = "data:image/png;base64,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"
$MediationGraphicBase64 = "data:image/png;base64,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"
$DBGraphicBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFIAAABVCAIAAABVSyR0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAgYSURBVHhe7ZttTFNnFMdBKIqE1xhcrBsiW1m0E9nCGrCAuAyMYSZoAiYWI2Yvblo/4NzYJsyISzCbfqnZ3JagQb/AIiTMLINkuGHd0rkAkk4sk7eFghaBAgJVYOwvz8OlvX2l3JZLyy+NnPPcprf/+5znnPP0Xn2np6d9vI8V9K+XsSzbm1hia7tNN/po5Knq/oB20NA9aEgWRRx+I5oemw+8lq1qG+wbftKuG23pedw9ON6iHaEHZjmWvvFYRgx15gNfZA+PT7b0jNzVDvc/nmjsGtIOjHcPjNNj1uFMtlarvXr1ilqtHh0dpUMcERQUJBaLZbJcoVBIh2aBwu1fKKkzH5yWbZLSOjo65PKjKpWKc80An4lPzs/Px1no0OJhIvvSpVJquYyJiaduOItdTGRrNK3UciXcnsUw+d+Bq83p39zW6OYRoXOydTodpoI6roTbs/zQ2FvR1FvfNpCqUHU5kAUJS75d2SIMIYZ+fCL7chMmn7i24a/skEB/ZGm81kcE0iFLxK0L3i1eS+w7PcPya3f145MnqjWK+i6EgLWrwGfZgmf1KX2jMHwVHbLEmdr71eqH1PHxuXJbi1fCCyEnqu9hwUvO/25xzc/Vbaztw4ffIzYDim10tDPdHwPKlXk5rKysotYMFus2JvnXz6Qw9n/9F9o1MsgCF2XYd/pMbRv1Z1jlv6L2/YQf/+77sq6duGWyuN3iSHKU4E//WgGaT58upo5TFBUVovmhjstIf3nNm6I1ccJg4oYGCop3vYSYr733CHH+brlaErVtbfBKchQs+ZQGKg7GV7/9mjwlKiUmgrxiI4MwXrY/jhjIdvJrLTPvpdgJcsx2Xt4h6jgFmhPztozDID+QEhUWaDVmq9W67MuNMMICBQ+Kd5BBYEe2i+BQtu2eHBH+3MlfSD7XfJoSNVsUPCHIbYB8tkrgR2y9YZIYwMNl/9mlx8ImNlnnhCUvG8KoZYmTP/1DjNejwjDzxAZLQPYmYbAkJpy80LrR0VlqNY/QkMG40zOCtox5oVGXnP8DBnnbh2km3ccSSGnGsNIb066gP8Fkoi2z2I3mJgi/zxFTZwYPWdvYhKAbKd4lor4R7yQ+z9IM+DvbLLZGha70X9HSMzI8PpeQMdupr0Seq+tASCNjqfKTsBVBT06Oxq0L+W6fGHsV4hqzZGa7qWsI4W2smQBVZbIthq8y7nwkRZwr9m5iMrZG9/jJxBSxWXhIkDNAORpV9GSwsc6xISHjLDxNNkArVrZ/CylXyRvDySALD5QNsCFDzHd9vh0GHTLFM2UDzLnxTpMFL2STH1L2JqyTvBiBdE1HXQkvChiLvpEn7box5sZQ+8NRjNBjptjdgVmDj7ItgupFbgM2/TuMq0BuA3q+bHPadKNYEbZ/V7XGEk5pMZFBzmkGHpvJbbMs25tYlu1NLMv2JrxUNn+7tI6OjtbW1s7OTq22mxlhbp6KxfTnsbCw8A0bNohEImbEEfglG6qUSuWtW0rnbpJGR0fHx7+anp4eGWlyW9ccvsiemJiorLyGcU6ebMnI2CmTyYKC5m6DsODL2lYoFOXl5Vw9zVNT83NRUSF1LMEX2UrlTWpxBBIBoI4ZfJEtlSZTiyOwzgF1zOCLbLlcnpOTIxAEUH9hYG3bfvaEFykN+UwgePbLtndlcpwaGSg7O2fHjrnnLLAyna7buGoVFeU2JpwvssmpMUsSiWTr1nhADs0LXJfGxkYEC0lm5mWSwY5sR55LC8MFD7d88wH09vYaDAbqzMKaB/NTY5HHxoqEwvVhYaEiUSxZAsyXwfvBzBt9EBFjY2OtrZrubq1eb/KIi/OyETkLfC7NESyeeuHYkM2LTO7r60std2FnthFUjjyXtnnzZvLVmQwcEBCANENsBq1WOzg4F4exsTR6Jycn8/IOmj+juUCcD3IHKSk5S0QaP2p57tx5Vl7AKfAG/EtcFJuCggKiHEkIh7hV7vIgv3Gjjhjbts09bVJbW0utWZCokSmYHUJjY0NJSQmKNmxcIONDroYb2egxyLeXSqVMp4X9gEqlIjYDT5RzIxvB2dDQAANfeufODDIILlxQoJBSZxaWvEVR7nfq1Cli4atfv36d2E4wNjaampoKAzWvubm5v78fNsTU1//W19c3OKhH9caqJmBcKBQ2NdEr8uBBb1tbW1JSkp+fH1oA9CpoOciFWAg5OfuoZYadlIaYTEtLo4498HXRWsDARyFXsZoHu3Ce4RahXdHr9QoFIvxZ5DsONmHMFNXV1WGNENs5FqFdQcdaWFh49KgcDSYdsgdme8+evcTGbLv0f8lxGeSEkJCQxMRE9OnUn8ka6JzRNlN/Bix4zCd1XFPDXd6usMC3z83Nzcx8i/pm4Fxu6FvcHeRIwqWlpQUFH9fU1DDaGCDp7NkSZhxFKz8/33W9mkVcMtsskOELCj4hwiAJwiCPHDIu1PgCx4/nc6h5EVKaMTJZriOa3TPPBHfM9sWL3yI1wkCv2tlJNQMsfqJ5amrqyJEPmLDnikWebQaJRIKyzLyYJhRFnnPNtnGrbGtgtqnlLuaCHOlXozEprVyxevVq1u8nWNXUmgGzfehQHnW4w6G6vbigbrP6mQWC3QE2dtQxgxdBDrKysqjFEZmZmdSyBF9kI9uhgedkpy0QBOCjpFL2/yoyhi9BTkDdrqqqUipvOpfYsRFITpZmZe3BRogOWYFfshnQ0qjV6qGhISx4XAJrVwHtAMAyDg0NxR6ZlSltwFPZroYva9vNLMv2Hnx8/ge19eXZ1S6HlQAAAABJRU5ErkJggg=="
$ReplicationGraphicBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEYAAABGCAIAAAD+THXTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAeySURBVGhD7Zt/TJVVGMfJSciaV6OAGioaFoxQuTlNWV6brtuWOTfjjwzdxM2w/sCmU/9INzfxH1ot7x8paws20/xD21zdNiFZ6oYYNSARf4HmDyjBYUEOTZd97/s+3s573vOe95z3vl7K9Rmb57lw73u+5zzPc57ngI/cu3cv5eFiFP37EPG/pP8CDySW6jv6OnsGMTjd+8fgrTuDw3dNMy/rscxAWtroUcWTxplmeFo2zNh7/MM3SZj3kTP9Daf6j5y5Pjh8h15VIFyUHcp/IjwtK3NsGr2UGD5Iwp7sPX4FSsj2SjB3/FtzJ5TOyiHbKwlJau4eqI6eb730G9l+AG/cuOi5cFEW2fp4lHR1YHjbwbP1HdfI9ps5UzPWhvPm5GWQrYMXSdXfnN95+AIZD5LFwac/eLNIN3/oSUIO2LCv48Ftjp3CnEDNyuIJGelkK6AhCc628tMfu/tukp0sAumpdatfQPIg2w1VScgEFbVtWtnZR+B7W5YUlJVMJFuKkiToWfZJCxkjB1StCuWS4Yx75MHfsD9kjCjV0XNYXDKccZGEfID4GSl/47h99y8sLpaYbAdcJFV+3p78fCABi1tR1wZtZIuQScJhmniZ4zuogCt3/0SGCMf00NkztOijJjI8gaKmOHd8MHdcrAA3SlJseP/QbThzQ0cfDjcMzJ/0QE150KlocpQEr/V2pAbSR5fNnVgeynWtrOECuxovqkS8HRzB0XVzybAiloSHISuQoUPVG4Wls3O0ShgU8khlHiIWtZKwbBc/G1UcjTS50H9TtySD/0TXl6CcI1sZNAHCPCF4fLwnVYET8NnRS/gigwEhJNkHfEhk+XScpGSrgc/c03SFDAaB4y2NnFBsgVD8V5UWLo00c4Fuxi4OELSGzd032E9DpJltrHBbMMXNBzrJUADh+v3Wl8m4Dy8J0mdv/Y4MKaiOo+tKMEXsakVtK71qgFVHlSmPe7wd22LPWpAkXHsnvqx8katoecdTPIgw6brVM6EHY0wLfaj5uglc3DWPGXVWKwRw8QCdhTljyVDAPmFe0tdtv9JISmTFDJw2ZKSkvLNgiof4BtgQCGNVYbGQNslQwD5hi6TY6na5nxLCuwGkVPWWhgXLvO3gGTIM8Dn2z3cCWYer+iyS6k9eE6ZFFjwMe0IGA1a3przY28UV9goBSYbBGtEjnODea5Hket7By+FyZNiAHqiCNrJ12LzfElTYKPXm/HTvEI0MLI/vuXGLRiKQDGpWBuUzxlTggWTogEz7VesvZBio+x7eSyMDy/y477EYfhVUWTnkCaFnuoJalkYGofwnaeRG/9CfNDLgJFm+x4Lcqn6r5u1ukUvH6o7XPyjZJev34pSVTFS8yojDZXkVEEts7lLPNJxzWaqHKesP0YgBq4WDIi3VIh4xEw8qVEO1xwR1Hea3v6WHDDW4UkA4HyHH3g/Fd/UfSVikgk0N5tiVbze9xG6C+rNNkDkDY1LJYOBKB/u11NUbw9wpZCKWBGZsblS8Ofni3VlsaM3bflT4JCfw3rq3Z+qme5wx9hLZpL1qoVmdAcuHZo59lEZucAImPK4ayiaoAHEQkaEGAsa4qxLowdLE9QCrpIB6RFpyY6hANeHGQZjtqO8mww0EheS6i5u2RZL6YnNnyOLip2ikw8eHurjj1YnK3e2SFo5L9x53CQ9g1wwfqpvlTTbs63BtQ1DUclUcB5fuLZJyxo+hkQLcY9aG81iHVsTwqFZJanHq/FlkuzRfJyR2NV5kC00sVWS5Y0UrARG/bGeLMO7tfYcQtP00MrDGUkY6e9LJQQri1g8r4q1mxS4hm7ELBDp7hrj+XwiWkqvULJIAp1jOjkNdnM+UzsqJN/BaIDgRV2Q4iBRi9yxekla/HYuEOkubDfAM1BbyXwThJEE64YpAZL/q6DkM4ITQI2kLWF63JVvBpZduKYBViCyfTgYD5lR/su/o2es4xMwUDA/BBqJrwFswwFOWRk5wU0c92XCqT/1Wp337Qq4KEUhChKgEJQum6OFX3wBSUcipOJiQVaHJW5bkk3EfwSTgEuqFvQl8BjNTdBUWZCNJ5y8HK7hmwWQyGASSjB/Vbkux3os+PO56hnBgf1Aoe0gnoKxkknDpBY4H8CR4ufrNOAtOgo2vPTu/IFM+UWzp3uNX9zRd8bC3AGKQhISPEEsC9mthXZD6Xnk+Ky87ltaQ3ALpqWaSaL30+4nugeauAc8hBNBZOSVVR0lA/b4/ycARsEVO2UiWo2rK9f7SJTkYf6wiax9lkuCvrhd3yQenhfyixmW68uvV5IN63/U6zX0HYpfgC58hY0QJF2W/9+pUMpxRciokZZzTZIwQ0BNZISi77MgyHgdKBBTLiWRez8DfVPbHREMSiPUwdbIm1HeQ35APtK6j9SQBHPYVtW3JOa9whCDlav2eE2hLMtnf0lMdPe+tllEBm1M+bxLSkocjxKMkgKDaefhC7bHL/v5pGzSgHvV2OWPiXZIJNmpX4897mi77kjbQ50NMgiVLopLioMxFA4s2Vtcb4WNGgZvp1/+/8E1SHGTFAz/0dvbGGhOM7W6JeQcnx+6hUNegVLffhySI/5LsDA7fNRVCg26/7IFkSEoyPvjuv42HTlJKyt/4sM4skLKz8QAAAABJRU5ErkJggg=="
#endregion



#region funtions
function CheckServerCSService {
[cmdletbinding()]
Param (
 [Parameter(Position=0,ValueFromPipeline=$True)]
 $Server
 )
 Process
 {
    $NoError = $true
    if (Test-Connection -computer $Server -quiet) {
        $csw = get-CsWindowsService -computername $Server -ExcludeActivityLevel
    
        foreach ($svc in $csw){
            if ($svc.status -notlike "Running") { $NoError = $false }
        
        }
    }else{ $NoError = $false }  # End If
    return $NoError
 } #Close Process
} #Close Function

function CheckServerResponds {
[cmdletbinding()]
Param (
 [Parameter(Position=0,ValueFromPipeline=$True)]
 $Server
 )
 Process
 {
    $NoError = $false
    if (Test-Connection -computer $Server -quiet) { $NoError = $true }
    return $NoError
 } #Close Process
} #Close Function

function BuildStatusTable($ServerList,$pingonly) {
    Foreach ($PoolIdentity in $ServerList.Identity) {
        $ArrComputer =@() # Create a blank Array For Shortened Computer NETBIOS name
        $ArrFQDNComputer =@() # Create a blank Array For FQDN name
        $FilteredPools = $ServerList | where { $_.identity -eq $PoolIdentity} #create a new object for each 'group' or servers within a pool
        Foreach ($Computer in $FilteredPools.Computers) {
            # Loop through the computers in each pool
            $strComputer = $Computer.ToString()
            $ArrFQDNComputer += $strComputer
            # Remove Trailing Domain name if there - everything after '.'
            $strComputer = $strComputer.Substring(0,$strComputer.IndexOf('.'))
            #add the filtered computer name to the array
            $ArrComputer += $strComputer
        }
        Write-verbose $poolIdentity 
        $CommonServerName = get-LongestCommonSubstringArray $ArrComputer
        $ReturnHTML += "<h3>$poolIdentity ( $CommonServerName )</h3><table><tr>"
        $arrFQDNcomputer | foreach { 
                Write-host "Checking Services on $_" -ForegroundColor DarkGray
                #
                # Lets check to see if we have already scanned this machine
                #

                if ($HashCheckedComputers.ContainsKey($_)) {
                    $ServerOK = $HashCheckedComputers.get_item($_)
                        If ($ServerOK) { 
                            $ComputerTDclass = "computer_pass"
                        } else {
                            $ComputerTDclass = "computer_fail"
                        }
                } else {
                    # Not in prescanned list - lets check the server is ok
                    if ($pingonly) {
                        $ServerOK = $_ | CheckServerResponds
                        If ($ServerOK) { 
                            $ComputerTDclass = "computer_pass"
                        } else {
                            $ComputerTDclass = "computer_fail"
                        }
                    } else {
                        $ServerOK = $_ | CheckServerCSService
                        If ($ServerOK) { 
                            $ComputerTDclass = "computer_pass"
                        } else {
                            $ComputerTDclass = "computer_fail"
                        }
                    $HashCheckedComputers.Add($_,$ServerOK)
                    }
                }
                $strHTMLComputer = $_.Substring(0,$_.IndexOf('.'))
                if ($CommonServerName -ne $null){ 
                    if ($overrideAutoStrip){
                        $strHTMLComputer = $_.Substring(0,$_.IndexOf('.'))
                        $strHTMLComputer = $strHTMLComputer.substring($strHTMLComputer.length - $charstokeepforserverid,$charstokeepforserverid)
                    } else {
                        $strHTMLComputer = ($_.Substring(0,$_.IndexOf('.'))).TrimStart($CommonServerName)
                    }
                }

                $ReturnHTML += "<td class='$ComputerTDclass'> $strHTMLComputer </td>"
            } 
            $ReturnHTML += "</tr></table>"
        $ArrComputer = $null
        

    } # end foreach
    return $ReturnHTML
} # end function


function CheckPing($server)
{
    if (Test-Connection -ComputerName $server -Quiet) {$pingresult = $true} else {$pingresult =$false}
    return $pingresult
}

function CheckReplicationStatus($ReplicaFqdn)
{
    $ReplicaStatus = (Get-CsManagementStoreReplicationStatus -ReplicaFqdn $ReplicaFqdn).uptodate
    return $ReplicaStatus
}

Function Convert-HTMLEscape 
{
 <#
 Convert &lt; and &gt; to < and >
 It is assumed that these will be in pairs
 Also convert $quot; to "
 #>
[cmdletbinding()]
Param (
 [Parameter(Position=0,ValueFromPipeline=$True)]
 [string[]]$Text
 )
  Process
  {
    foreach ($item in $Text)
    {
      if ($item -match "&lt;BREAK-NL&gt;")
      {
        $item.Replace("&lt;BREAK-NL&gt;","<BR>")
      }
      else
      {
        #otherwise just write the line to the pipeline
        $item
      }
    }
  } #close process
} #close function


Function get-LongestCommonSubstringArray
{
Param(
[Parameter(Position=0, Mandatory=$True)][Array]$Array
)
    $PreviousSubString = $Null
    $LongestCommonSubstring = $Null
    foreach($SubString in $Array)
    {
        if($LongestCommonSubstring)
        {
            $LongestCommonSubstring = get-LongestCommonSubstring $SubString $LongestCommonSubstring
            write-verbose "Consequtive diff: $LongestCommonSubstring"
        }else{
            if($PreviousSubString)
            {
                $LongestCommonSubstring = get-LongestCommonSubstring $SubString $PreviousSubString
                write-verbose "first one diff: $LongestCommonSubstring"
            }else{
                $PreviousSubString = $SubString
                write-verbose "No PreviousSubstring yet, setting it to: $PreviousSubString"
            }
        }
    }
    Return $LongestCommonSubstring
}

Function Add-HTMLTableAttribute
{
    Param
    (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $HTML,

        [Parameter(Mandatory=$true)]
        [string]
        $AttributeName,

        [Parameter(Mandatory=$true)]
        [string]
        $Value

    )

    $xml=[xml]$HTML
    $attr=$xml.CreateAttribute($AttributeName)
    $attr.Value=$Value
    $xml.table.Attributes.Append($attr) | Out-Null
    Return ($xml.OuterXML | out-string)
}    

Function get-LongestCommonSubstring
{
Param(
[string]$String1, 
[string]$String2
)
    
    if((!$String1) -or (!$String2)){Break}

    
    # .Net Two dimensional Array:
    $Num = New-Object 'object[,]' $String1.Length, $String2.Length
    [int]$maxlen = 0
    [int]$lastSubsBegin = 0
    $sequenceBuilder = New-Object -TypeName "System.Text.StringBuilder"

    for ([int]$i = 0; $i -lt $String1.Length; $i++)
    {
        for ([int]$j = 0; $j -lt $String2.Length; $j++)
        {
            if ($String1[$i] -ne $String2[$j])
            {
                    $Num[$i, $j] = 0
            }else{
                if (($i -eq 0) -or ($j -eq 0))
                {
                        $Num[$i, $j] = 1
                }else{
                        $Num[$i, $j] = 1 + $Num[($i - 1), ($j - 1)]
                }
                if ($Num[$i, $j] -gt $maxlen)
                {
                    $maxlen = $Num[$i, $j]
                    [int]$thisSubsBegin = $i - $Num[$i, $j] + 1
                    if($lastSubsBegin -eq $thisSubsBegin)
                    {#if the current LCS is the same as the last time this block ran
                            [void]$sequenceBuilder.Append($String1[$i]);
                    }else{ #this block resets the string builder if a different LCS is found
                        $lastSubsBegin = $thisSubsBegin
                        $sequenceBuilder.Length = 0 #clear it
                        [void]$sequenceBuilder.Append($String1.Substring($lastSubsBegin, (($i + 1) - $lastSubsBegin)))
                    }
                }
            }
        }
    }
    return $sequenceBuilder.ToString()
}


function CheckServicesAreRunning{
[cmdletbinding()]
Param (
 [Parameter(Position=0,ValueFromPipeline=$True)]
 $Server
)
Process
 {
 $CheckServer = Get-CsWindowsService -ComputerName $Server -ExcludeActivityLevel
 $anError = $True
 foreach ($svc in $CheckServer)
    {
    if ($svc.status -notlike "Running") {$anError = $false}
    }
    $anError
 } #close process
} #close function


function CheckCSDatabaseServicesAreRunning{
[cmdletbinding()]
Param (
 [Parameter(Position=0,ValueFromPipeline=$True)]
 $PoolIdentity
)
Process
 {
 $IdentityToCheck = "UserDatabase:$PoolIdentity"
 $NoError = $False
 if (Test-Connection -ComputerName $PoolIdentity -Quiet)
  {
    if ((Get-CsUserDatabaseState -Identity $IdentityToCheck).online -notlike "True") {$NoError = $False} else {$NoError = $True}
  }
  return $NoError
 } #close process
} #close function

Function BuildDatabaseTable()
{
    $ReturnHTML = "<table><tr><th>SQL Pool</th><th>Online</th></tr>"
    foreach ($dbpool in $DatabaseServers.identity){
        $ServerOK = $dbpool | CheckCSDatabaseServicesAreRunning
        If ($ServerOK) { 
            $ComputerTDclass = "computer_pass"
        } else {
            $ComputerTDclass = "computer_fail"
        }
        $strHTMLComputer = $dbpool.Substring(0,$dbpool.IndexOf('.'))
        $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$ServerOk</td></tr>"
    } # end for each
    $ReturnHTML += "</table>"
    return $ReturnHTML
} #end function


function BuildCSIMTable(){
    $ReturnHTML = "<table><tr><th>Pool</th><th>Completed</th><th>Latency</th></tr>"
    $PoolsWithTestAccounts = Get-CsHealthMonitoringConfiguration
    foreach ($dbpool in $PoolsWithTestAccounts.identity){
        $ServerOK = Test-CsIM -TargetFqdn $dbpool
        If ($ServerOK.Result -like "Success") { 
            $ComputerTDclass = "computer_pass"
        } else {
            $ComputerTDclass = "computer_fail"
        }
        $strHTMLComputer = $dbpool
        $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$($ServerOk.result)</td><td>$($ServerOk.Latency)</td></tr>"
    } # end for each
    $ReturnHTML += "</table>"
    return $ReturnHTML

}
#endregion


#region Main Program

 
#region Verifying Administrator Elevation 
Write-Host Verifying User permissions... -ForegroundColor Yellow -NoNewline 
#Start-Sleep -Seconds 1 
#Verify if the Script is running under Admin privileges 
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(` 
  [Security.Principal.WindowsBuiltInRole] "Administrator"))  
{ 
  Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!" 
  Write-Host  
  Break 
} 
else 
{ 
    Write-Host " Done!" -ForegroundColor Green 
} 
#endregion 
 
#region Import Lync / SfB Module 
Write-Host 
Write-Host "Please wait while we're loading Skype for Business / Lync PowerShell Modules..." -ForegroundColor Yellow -NoNewline 
  if(-not ((Get-Module -Name "Lync")) -or (Get-Module -Name "SkypeforBusiness")){ 
    if(Get-Module -Name "Lync" -ListAvailable){ 
      Import-Module -Name "Lync"; 
    } 
    elseif (Get-Module -Name "SkypeforBusiness" -ListAvailable){ 
      Import-Module -Name "SkypeforBusiness"; 
    } 
    else{ 
      Write-Host Lync/SfB Modules are not exist on this computer, please verify the Lync Admin tools installed -ForegroundColor Red 
      exit;    
    }     
  } 
Write-Host -NoNewline     
Write-Host " Done!" -ForegroundColor Green 
#endregion 


#region Collating pools and servers
$Pools = Get-CsPool | select Identity,Services | sort -Property Identity


    Write-host "Finding All Pools & Services in environment.." -ForegroundColor Yellow -NoNewline 
    Foreach ($pool in $Pools)
        {
    
        Write-verbose "[ ] $($Pool.Identity)" 
        $PoolServices = $pool.Services
        Foreach ($service in $PoolServices)
            {
            Write-verbose " - [ ] $service" 

            }
        }
    
    Write-Host " Done!" -ForegroundColor Green 



$RegistrarServers = Get-CsPool | Where-Object {$_.Services -like "Registrar:*"} | Select identity,computers
$MediationServers = Get-CsPool | Where-Object {$_.Services -like "MediationServer:*"} | Select identity,computers
$EdgeServers = Get-CsPool | Where-Object {$_.Services -like "EdgeServer:*"} | Select identity,computers
$DatabaseServers = Get-CsPool | Where-Object {$_.Services -like "UserDatabase:*"} | Select identity,computers


#endregion

#region Building the HTML

$PoolsServicesHTML = "<H3>Pools and Services <a href='#poolinfo' class='nav-tab poolinfo'>[+]</a></h3><div id='poolinfo' class='poolinfo'><h5 class='subtitle'>Services and pools within each pool, use this table to identify what each pools role is</h5>"
$PoolsTableWithAttib = ($replaceBreak = $pools | Select Identity,@{Name="Services";Expression={[string]::Join("<BREAK-NL>",($_.Services))}} | ConvertTo-Html -Fragment | Out-String | Add-HTMLTableAttribute -AttributeName "class" -Value "pools").replace("&lt;BREAK-NL&gt;","<br>")
$PoolsServicesHTML += $PoolsTableWithAttib
$PoolsServicesHTML += "</div>"


Write-host "Checking Servers in the environment.." -ForegroundColor Yellow
    





$CSWindowsServicesHTML = "<H3>Individual Server Check</h3>"
$CSWindowsServicesHTML += "<h5 class='subtitle'>Dashboad showing each server sorted by type then pool, use this table to identify if an individual server is running all windows services and/or responding to pings</h5>"

$HashCheckedComputers =@{}
#Checking Registar Servers
$CSWindowsServicesHTML += "<div class='FEandMed_Container'>"
$CSWindowsServicesHTML += "<div class='servers'><h2 class='divheader'>Registrar Servers</h2><img class='servergraphic' src='$FrondEndGraphicBase64'><div class='serverdivcontent'>"
$CSWindowsServicesHTML += BuildStatusTable $RegistrarServers $false
$CSWindowsServicesHTML += "</div></div>"

#Checking Mediation Servers
$CSWindowsServicesHTML += "<div class='servers'><h2 class='divheader'>Mediation Servers</h2><img class='servergraphic' src='$MediationGraphicBase64'><div class='serverdivcontent'>"
$CSWindowsServicesHTML += BuildStatusTable $MediationServers $false
$CSWindowsServicesHTML += "</div></div>"
$CSWindowsServicesHTML += "</div>"

#Checking Edge Servers
$CSWindowsServicesHTML += "<div class='EdgeandSQL_Container'>"
$CSWindowsServicesHTML += "<div class='servers'><h2 class='divheader'>Edge Servers</h2><img class='servergraphic' src='$EdgeGraphicBase64'><div class='serverdivcontent'>"
$CSWindowsServicesHTML += BuildStatusTable $EdgeServers $true 
$CSWindowsServicesHTML += "</div></div>"

#Checking SQL Servers Respond to Ping
$CSWindowsServicesHTML += "<div class='servers'><h2 class='divheader'>Database Servers</h2><img class='servergraphic' src='$DBGraphicBase64'><div class='serverdivcontent'>"
$CSWindowsServicesHTML += BuildStatusTable $DatabaseServers $true 
$CSWindowsServicesHTML += "</div></div>"
$CSWindowsServicesHTML += "</div>"


#Lets the Hash table
$HashCheckedComputers.Clear()

Write-Host " Done!" -ForegroundColor Green 
Write-Host
Write-host "Checking Services in the environment.." -ForegroundColor Yellow -NoNewline 

$ReplicationStatusHTML = "<H3>Lync Services Check</h3>"
$ReplicationStatusHTML += "<h5 class='subtitle'>Dashboad showing the status of replication and database services</h5>"

$ReplicationStatusHTML += "<div class='Replication_Container'>"

$ReplicationStatusHTML += "<div class='repservices'><h2 class='divheader'>CMS Management Replication Status</h2><img class='servergraphic' src='$ReplicationGraphicBase64'><div class='serverdivcontent'>"
$ReplicationStatusHTML += Get-CsManagementStoreReplicationStatus | Select ReplicaFqdn,uptodate,LastStatusReport,LastUpdateCreation,ProductVersion | ConvertTo-Html -Fragment
$ReplicationStatusHTML = $ReplicationStatusHTML.replace("<td>True</td>","<td class='replica_pass'>True</td>")
$ReplicationStatusHTML = $ReplicationStatusHTML.replace("<td>False</td>","<td class='replica_fail'>False</td>")
$ReplicationStatusHTML += "</div></div>"

$ReplicationStatusHTML += "<div class='sqlservices'><h2 class='divheader'>User Database Online</h2><img class='servergraphic' src='$ReplicationGraphicBase64'><div class='serverdivcontent'>"
$ReplicationStatusHTML += BuildDatabaseTable
$ReplicationStatusHTML += "</div></div>"

$ReplicationStatusHTML += "</div>"

Write-Host " Done!" -ForegroundColor Green 

Write-Host
Write-host "Checking Communications within the environment.." -ForegroundColor Yellow -NoNewline 

$HealthHTML = "<H3>Lync Communications Check</h3>"
$HealthHTML += "<h5 class='subtitle'>Are communications between the two test accounts working?</h5>"

$HealthHTML += "<div class='connectivity_Container'>"

$HealthHTML += "<div class='servers'><h2 class='divheader'>Health Monitoring Config</h2><img class='servergraphic' src='$ReplicationGraphicBase64'><div class='serverdivcontent'>"
$HealthHTML += Get-CsHealthMonitoringConfiguration | Select Identity,FirstTestUserSipUri,FirstTestSamAccountName,SecondTestUserSipUri,SecondTestSamAccountName,TargetFqdn  | ConvertTo-Html -Fragment -as List
$HealthHTML += "</div></div>"

$HealthHTML += "<div class='servers'><h2 class='divheader'>IM between test users</h2><img class='servergraphic' src='$ReplicationGraphicBase64'><div class='serverdivcontent'>"
$HealthHTML += BuildCSIMTable
$HealthHTML += "</div></div>"
  

$HealthHTML += "</div>"

Write-Host " Done!" -ForegroundColor Green 


Write-Host
Write-host "Building HTML Report.." -ForegroundColor Yellow -NoNewline 


#get-LongestCommonSubstring "BigMan-LyncFE1" "BigMan-LyncFE2"

$DivandImageHTML = "<div class='entirepage'><img class='headerimage' src='$CompanyLogoBase64'>"

ConvertTo-Html -Head $Header -Body $DivandImageHTML,$PoolsServicesHTML,$CSWindowsServicesHTML,$ReplicationStatusHTML,$HealthHTML -PostContent "<br> </div><i>Reported Created $(get-date)</i>" | Out-File PShtml.html
Invoke-Item PShtml.html

#endregion

Write-Host " Done!" -ForegroundColor Green 
#endregion
