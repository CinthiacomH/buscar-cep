Sub busca_cep()

Range("B3:D3").ClearContents

'Abrindo o navegador'
Set ie = CreateObject("internetexplorer.application")

'Abrindo o endereço dos correios'
ie.navigate "http://www.buscacep.correios.com.br/sistemas/buscacep/buscaCepEndereco.cfm"
'Informar se deseja que o navegador seja aberto como um popup'
ie.Visible = True

'Aguardar carregamento total da página web'
Do While ie.busy And ie.readyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

'Capturando valor e selecionando botão'
ie.document.getElementsByTagName("input")(0).Value = Cells(3, 1).Value
ie.document.getElementsByClassName("btn2 float-right")(0).Click
    
'Aguardar carregamento total da página web'
Do While ie.busy And ie.readyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

'Preenchendo as colunas B3, C3 E D3'
Cells(3, 2) = ie.document.getElementsByTagName("td")(0).innertext
Cells(3, 3) = ie.document.getElementsByTagName("td")(1).innertext
Cells(3, 4) = ie.document.getElementsByTagName("td")(2).innertext

'Fechar o mavegador'
ie.Quit

Range("A3:D3").WrapText = False

End Sub
