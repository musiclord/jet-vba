# 大標題

### 小標題

### 應用場景

使用者為一般審計員及中央小組，審計案件的資料量應不超過一百萬筆，

### 技術規格
- VBA
    - VBA 的類別模型是比較簡化的 `COM` 架構，與 C#、Java 等語言不同，不支援多個建構子（Overloaded Constructors），不允許 Class_Initialize() 帶參數，無法在 New 關鍵字後傳入參數
    - 要能夠依賴注入，得自定義方法:
    ```vb
    ' 類別模組: clsCustomer
    Private m_name As String
    Private Sub Class_Initialize()
        ' 預設初始化
    End Sub
    Public Sub Initialize(ByVal p_name As String)
        m_name = p_name
    End Sub
    ' 類別模組: 主程式
    Dim customer As clsCustomer
    Set customer = New clsCustomer
    customer.Initialize("Alice")
    ```

- Access
    - 

- Excel
    - 

### 開發架構

- MVC
    - Model: 
    - View: 
    - Controller: 

- Three-tier Layer

### 引用項目

- DAO:
    - Name: `Microsoft Office 16.0 Access database engine Object Library`
    - 

- ADO:
    - Name: `Microsoft ActiveX Data Objects 6.1 Library`
    -

- ADOX:
    - Name: `Microsoft ADO Ext. 6.0 for DDL and Security`
    -

- Script:
    - Name: `Microsoft Scripting Runtime`
    - 