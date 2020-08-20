
<%
'---Ajusta los parametros con los de tu servidor:
    Ip= "172.16.91.2"    
    Instance = "DCCENTRAL"         
    DbPort="1433"
    DbName ="DcEstatal"
    DbUser = "itv_dba"
    DbPassword = "818t3m41t4vu"
    Token = "it4vu" 'Esta es la contraseña para que accedan al webservice
'-----------------------------------------------------------------------
'Recomendacion: 
' El usuario que pongas en este modulo, 
' sugiero que solo tenga accesa acceso 
' a consultas con select, a menos que 
' requieras otra cosa
'----------------------------------------------------------------------
'
'
' El webservice recibe:
' url/sqlservertojson.asp?method=GET&token=MiToken&sql=
'
' 
'
'         Cualquier duda o sugerencia
'          Estoy a tus ordenes:
'Juan Jose Pedraza      | printepolis@gmail.com
'WhatsApp: 8343088602   | facebook.com/prymecode
'========================================================================
    
    Mode = Request.QueryString("method")
    if Mode = "GET" then       
        SQLsolicitado = Request.QueryString("sql")
        ElToken = Request.QueryString("token")
    else        
        SQLsolicitado = Request.Form("sql")
        ElToken = Request.Form("token")
    end if

    if (Token = ElToken) then
            CadenaDeConeccion = "PROVIDER=SQLOLEDB;DATA SOURCE=" & Ip & "\" & Instance & ";UID=" & DbUser & ";PWD=" & DbPassword & ";DATABASE=" & DbName
            Set cnnSolicitada = Server.CreateObject("ADODB.Connection")                    
            cnnSolicitada.Errors.Clear()
            cnnSolicitada.ConnectionTimeout = 0                    
            cnnSolicitada.CommandTimeout = 0 
            cnnSolicitada.open CadenaDeConeccion        

                set rsSolcitada=cnnSolicitada.Execute(SQLsolicitado) 
                if not rsSolcitada.eof then
                    ' Response.ContentType = "text/html"
                    Response.CodePage = 65001
                    Response.CharSet = "UTF-8"                                           
                    Response.ContentType = "application/json"                         
                    response.Write "" & JSONData(rsSolcitada, "") & "" 
                else
                    Response.CodePage = 65001
                    Response.CharSet = "UTF-8"                            
                    response.ContentType = "application/json"
                    data = ""
                    data = data & "[{"
                    data = data &  """" & "r" & """" & ":""" & "Sin resultados"  & """"
                    data = data & "}]"
                    response.Write (data)
                end if


            cnnSolicitada.Close()  
        else
                Response.CodePage = 65001
                Response.CharSet = "UTF-8"                            
                response.ContentType = "application/json"
                data = ""
                data = data & "[{"
                data = data &  """" & "r" & """" & ":""" & "Token No Valido"  & """"
                data = data & "}]"
                response.Write (data)

        end if



Function JSONData(ByVal rs, ByVal labelName) 
		Dim data, columnCount, colIndex, rowIndex, rowCount, rsArray
		If Not rs.EOF Then
			data = labelName & "["
			rsArray = rs.GetRows() 
			rowIndex = 0
		End If
			rowCount = ubound(rsArray,2)
			columnCount = ubound(rsArray,1)
			For rowIndex = 0 to rowCount
				data = data & "{"
			   For colIndex = 0 to columnCount
                IF IsNull(rs.Fields(colIndex).Name) = False THEN
                    IF IsNull(rsArray(colIndex,rowIndex)) = False THEN
					
                        data = data &  """" & QuitaLoIndeseable(rs.Fields(colIndex).Name) & """" & ":""" & QuitaLoIndeseable(rsArray(colIndex,rowIndex)) & """"
                    else 
                        data = data &  """" & "" & """" & ":""" & "" & """"

                    End If
                else
                        data = data &  """" & "" & """" & ":""" & "" & """"

                End If
					If colIndex < columnCount Then
						data = data & ","
					End If
			   Next 
			   data = data & "}"
			   If rowIndex < rowCount Then
					data = data & ","
			   End If
			Next 
			data = data & "]"
			rs.Close
		JSONData = data
 End Function



Function QuitaLoIndeseable(Cadenilla)
    Cadenilla =  Replace(Cadenilla, Chr(34), " ") 'Comillas dobles
    Cadenilla =  Replace(Cadenilla, Chr(39), " ") 'Comillas simples
    Cadenilla =  Replace(Cadenilla, Chr(13), " ") 'saldo de carro
    Cadenilla =  Replace(Cadenilla, Chr(32), " ") 'saldo de carro
    Cadenilla =  Replace(Cadenilla, "-", " ") '
    Cadenilla =  Replace(Cadenilla, "#", " ") '
    QuitaLoIndeseable =  Cadenilla
    

End Function



%>
