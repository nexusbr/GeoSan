VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsX_Suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private cmd As ADODB.Command
Private params As ADODB.Parameters
'local variable(s) to hold property value(s)
Private mvarSupplierID As Long 'local copy
Private mvarCompanyName As String 'local copy
Private mvarContactName As String 'local copy
Private mvarContactTitle As String 'local copy
Private mvarAddress As String 'local copy
Private mvarCity As String 'local copy
Private mvarRegion As String 'local copy
Private mvarPostalCode As String 'local copy
Private mvarFax As String 'local copy
Private mvarHomePage As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCountry As String 'local copy
Private mvarPhone As String 'local copy

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String
Dim m As String
     






Public Property Let Phone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Phone = 5
    mvarPhone = vData
End Property


Public Property Get Phone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Phone
    Phone = mvarPhone
End Property



Public Property Let Country(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Country = 5
    mvarCountry = vData
End Property


Public Property Get Country() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Country
    Country = mvarCountry
End Property





Public Property Let HomePage(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HomePage = 5
    mvarHomePage = vData
End Property


Public Property Get HomePage() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HomePage
    HomePage = mvarHomePage
End Property



Public Property Let Fax(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fax = 5
    mvarFax = vData
End Property


Public Property Get Fax() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fax
    Fax = mvarFax
End Property



Public Property Let PostalCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostalCode = 5
    mvarPostalCode = vData
End Property


Public Property Get PostalCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostalCode
    PostalCode = mvarPostalCode
End Property



Public Property Let Region(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Region = 5
    mvarRegion = vData
End Property


Public Property Get Region() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Region
    Region = mvarRegion
End Property



Public Property Let City(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.City = 5
    mvarCity = vData
End Property


Public Property Get City() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.City
    City = mvarCity
End Property



Public Property Let Address(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address = 5
    mvarAddress = vData
End Property


Public Property Get Address() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address
    Address = mvarAddress
End Property



Public Property Let ContactTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactTitle = 5
    mvarContactTitle = vData
End Property


Public Property Get ContactTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactTitle
    ContactTitle = mvarContactTitle
End Property



Public Property Let ContactName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ContactName = 5
    mvarContactName = vData
End Property


Public Property Get ContactName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ContactName
    ContactName = mvarContactName
End Property



Public Property Let CompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyName = 5
    mvarCompanyName = vData
End Property


Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyName
    CompanyName = mvarCompanyName
End Property



Public Property Let SupplierID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Suppliers = 5
    mvarSupplierID = vData
End Property


Public Property Get SupplierID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Suppliers
    SupplierID = mvarSupplierID
End Property

Public Function InsertData(Conn As ADODB.connection) As Boolean
     'Código inserido em 15/05/2007 por ®viviani em substituição das linhas comentadas abaixo para eliminação de procedures
    Dim rs As New ADODB.Recordset
    Dim strExecute As String
    Dim cc As String
    Dim bb As String
    cc = "SUPPLIERID"
    bb = "X_SUPPLIERS"
    strExecute = X_SuppliersNewUpd()
    Conn.execute strExecute
    If frmCanvas.TipoConexao <> 4 Then
    rs.Open "SELECT max(SupplierID) from X_Suppliers", Conn
    Else
    rs.Open "SELECT max(" + """" + cc + """" + ") from " + """" + bb + """" + "", Conn
    End If
    SupplierID = IIf(IsNull(rs), 0, rs(0).value)
    rs.Close
'   ' Create connection and command objects
'   Set cmd = New ADODB.Command
'   With cmd
'       Set .ActiveConnection = Conn
'       'Ocoorência de Procedure
'          'X_SuppliersNewUpd
'       .CommandText = "X_SuppliersNewUpd"
'       .CommandType = adCmdStoredProc
'       Set params = .Parameters
'   End With
'   params.Refresh
'   ' Specify input parameter values
'   params("@SupplierID") = SupplierID
'   params("@CompanyName") = CompanyName
'   params("@ContactName") = ContactName
'   params("@ContactTitle") = ContactTitle
'   params("@Address") = Address
'   params("@City") = City
'   params("@Region") = Region
'   params("@PostalCode") = PostalCode
'   params("@fax") = Fax
'   params("@HomePage") = HomePage
'   params("@Phone") = Phone
'   params("@country") = Country
'   ' Execute the command
'   cmd.execute , , adExecuteNoRecords
'   ' Retrieve stored procedure return value and output parameters
'   SupplierID = params(0)
'   Set params = Nothing
'   Set cmd = Nothing
'   Exit Function
End Function

Public Function UpdateData(Conn As ADODB.connection, Optional Upd As Boolean = False) As Boolean
   If Upd Then
        Dim strExecute As String
        strExecute = X_SuppliersNewUpd()
        Conn.execute strExecute
'      ' Create connection and command objects
'      Set cmd = New ADODB.Command
'      With cmd
'          Set .ActiveConnection = Conn
'          'Ocoorência de Procedure
'          'X_SuppliersNewUpd
'          .CommandText = "X_SuppliersNewUpd"
'          .CommandType = adCmdStoredProc
'          Set params = .Parameters
'      End With
'      params.Refresh
'      ' Specify input parameter values
'      params("@SupplierID") = SupplierID
'      params("@CompanyName") = CompanyName
'      params("@ContactName") = ContactName
'      params("@ContactTitle") = ContactTitle
'      params("@Address") = Address
'      params("@City") = City
'      params("@Region") = Region
'      params("@PostalCode") = PostalCode
'      params("@fax") = Fax
'      params("@HomePage") = HomePage
'      params("@Phone") = Phone
'      params("@country") = Country
'      ' Execute the command
'      cmd.execute , , adExecuteNoRecords
   Else
      Dim rs As ADODB.Recordset
       If frmCanvas.TipoConexao <> 4 Then
      Set rs = Conn.execute("SELECT * from X_Suppliers where Supplierid=" & SupplierID)
      Else
      a = "X_SUPPLIERS"
      b = "SUPPLIERID"
      Set rs = Conn.execute("SELECT * from " + """" + a + """" + " where " + """" + b + """" + "='" & SupplierID & "'")
      End If
      CompanyName = rs.Fields("CompanyName")
      ContactName = IIf(IsNull(rs.Fields("ContactName")), "", rs.Fields("ContactName"))
      ContactTitle = IIf(IsNull(rs.Fields("ContactTitle")), "", rs.Fields("ContactTitle"))
      Address = IIf(IsNull(rs.Fields("Address")), "", rs.Fields("Address"))
      City = IIf(IsNull(rs.Fields("City")), "", rs.Fields("City"))
      Region = IIf(IsNull(rs.Fields("Region")), "", rs.Fields("Region"))
      PostalCode = IIf(IsNull(rs.Fields("PostalCode")), "", rs.Fields("PostalCode"))
      Fax = IIf(IsNull(rs.Fields("Fax")), "", rs.Fields("Fax"))
      HomePage = IIf(IsNull(rs.Fields("HomePage")), "", rs.Fields("HomePage"))
      Phone = IIf(IsNull(rs.Fields("Phone")), "", rs.Fields("Phone"))
      Country = IIf(IsNull(rs.Fields("Country")), "", rs.Fields("Country"))
      rs.Close
      Set rs = Nothing
   End If
'   Set params = Nothing
'   Set cmd = Nothing
End Function

Public Function X_SuppliersNewUpd() As String
Dim str As String


a = "X_SUPPLIERS"
b = "COMPANYNAME"
c = "CONTACTNAME"
d = "CONTACTTITLE"
e = "ADDRESS"
f = "CITY"
g = "REGION"
h = "POSTALCODE"
i = "COUNTRY"
j = "PHONE"
k = "FAX"
l = "HOMEPAGE"


     If frmCanvas.TipoConexao <> 4 Then
         
    If SupplierID = 0 Then
        str = "Insert into X_Suppliers ( CompanyName, ContactName, ContactTitle, Address, City, Region, PostalCode, Country, Phone, Fax, HomePage) "
        str = str & "Values ('" & CompanyName & "','"
        str = str & ContactName & "','"
        str = str & ContactTitle & " ','"
        str = str & Address & "','"
        str = str & City & "','"
        str = str & Region & "','"
        str = str & PostalCode & "','"
        str = str & Country & "','"
        str = str & Phone & "','"
        str = str & Fax & "','"
        str = str & HomePage & "')"
    Else
    
    
     str = "Update X_Suppliers Set  CompanyName='" & CompanyName & "', "
        str = str & "ContactName='" & ContactName & "', "
        str = str & "ContactTitle='" & ContactTitle & "', "
        str = str & "Address='" & Address & "', "
         str = str & "City='" & City & "', "
   
        str = str & "Region='" & Region & "', "
        str = str & "PostalCode='" & PostalCode & "', "
        str = str & "Country='" & Country & "', "
        str = str & "Phone='" & Phone & "', "
        str = str & "Fax='" & Fax & "', "
        str = str & "HomePage='" & HomePage & "' "
        str = str & "where SupplierID='" & SupplierID & "'"
        

    End If
     
     Else
     a = "X_SUPPLIERS"
b = "COMPANYNAME"
c = "CONTACTNAME"
d = "CONTACTTITLE"
e = "ADDRESS"
f = "CITY"
g = "REGION"
h = "POSTALCODE"
i = "COUNTRY"
j = "PHONE"
k = "FAX"
l = "HOMEPAGE"
m = "SUPPLIERID"


      If SupplierID = 0 Then
        str = "Insert into " + """" + a + """" + " ( " + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + "," + """" + g + """" + "," + """" + h + """" + "," + """" + i + """" + "," + """" + j + """" + "," + """" + k + """" + "," + """" + l + """" + ") "
        str = str & "Values ('" & CompanyName & "','"
        str = str & ContactName & "','"
        str = str & ContactTitle & " ','"
        str = str & Address & "','"
        str = str & City & "','"
        str = str & Region & "','"
        str = str & PostalCode & "','"
        str = str & Country & "','"
        str = str & Phone & "','"
        str = str & Fax & "','"
        str = str & HomePage & "')"
    Else
a = "X_SUPPLIERS"
b = "COMPANYNAME"
c = "CONTACTNAME"
d = "CONTACTTITLE"
e = "ADDRESS"
f = "CITY"
g = "REGION"
h = "POSTALCODE"
i = "COUNTRY"
j = "PHONE"
k = "FAX"
l = "HOMEPAGE"
m = "SUPPLIERID"
        str = "Update " + """" + a + """" + " Set  " + """" + b + """" + "='" & CompanyName & "', "
        str = str + """" + c + """" + "='" & ContactName & "', "
        str = str + """" + d + """" + "='" & ContactTitle & "', "
        str = str + """" + e + """" + "='" & Address & "', "
             str = str + """" + f + """" + "='" & City & "', "
        str = str + """" + g + """" + "='" & Region & "', "
        str = str + """" + h + """" + "='" & PostalCode & "', "
        str = str + """" + i + """" + "='" & Country & "', "
        str = str + """" + j + """" + "='" & Phone & "', "
        str = str + """" + k + """" + "='" & Fax & "', "
        str = str + """" + l + """" + "='" & HomePage & "' "
        str = str & "where " + """" + m + """" + "='" & SupplierID & "'"
       
    End If
     End If
'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"

    
    X_SuppliersNewUpd = str
End Function
