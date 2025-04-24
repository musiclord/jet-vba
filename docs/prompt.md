**�W�U��:**
�n���A�o�O�ھڧڭ̥ثe��ܤ��e�A��M�׭��c�����A�B�ݨD�M�`�N�ƶ������I�`���G
1.  **�ؼЬ[�c:** �M�ת��֤ߥؼЬO�N VBA �{���X���c���@�Ӥ��h�[�c�A�]�t�G
    *   **View:** `vMain.frm`, `vMapping.frm` (�ϥΪ̤������)�C
    *   **Controller:** `cApplication.cls`, `cMapping.cls` (�B�z UI �ƥ�A��լy�{)�C
    *   **Service Layer:** `ImportService.cls`, `PreviewService.cls`, `GLService.cls`, `TBService.cls`, `MappingService.cls`, `FilterService.cls` (�ʸ˷~���޿�)�C
    *   **Data Access Layer (DAL):** `AccessDAL.cls` (�ʸ˻P Access ��Ʈw������)�C
    *   **Entities (Optional):** �p `GLEntity.cls`, `TBEntity.cls`, `FilterCriteria.cls` (�Ω��ƶǻ��Ωw�q���c)�C
    *   **Utilities:** `mod_Utility.bas` (�q�λ��U���)�C

2.  **��e���c�J�I:** �ثe���u�@�D�n�����b `cApplication.cls`�A�T�O���ŦX�����¾�d�G�ȳB�z�Ӧ� `vMain` ���ƥ�A������ε{�����D�n�y�{�]�p�פJ�B�w���B�B�z�^�A�ñN����u�@�e���������� Service �h�C

3.  **`cApplication` ¾�d����:** `cApplication` �Y���u���`�I������h�G
    *   **���]�t**�����ե� `AccessDAL` ���{���X�C
    *   **���]�t**��������ƳB�z�η~�ȳW�h�p��C
    *   ���k�]�p `ImportCSV`, `PreviewTable`, `PopulateTableList`�^�D�n�t�d�G�����ƥ�B�ǳưѼơB�ե� Service �h��k�B�B�z Service ��^���G�A�H�Χ�s UI ���A�]�p���A�C�B�T���ءB�ҥ�/�T�α���^�C

4.  **�ƥ��X�ʬy�{:** `vMain.frm` �ϥ� `Public Event` (�p `DoImportGL`, `GetTableNames`) ���n���ϥΪ̾ާ@�C`cApplication.cls` �ϥ� `Private WithEvents vMain As vMain` �Ӻ�ť�o�Ǩƥ�A�æb������ `vMain_EventName()` �B�z�{�Ǥ��A�z�L²�檺 `Call Me.PrivateSubName(...)` �y�y�եΤ����p����k���T���C`vMapping.frm` �M `cMapping.cls` �]�N��`�����Ҧ��C

5.  **��Ʈw�B�z (`AccessDAL.cls`):**
    *   �ϥΫ���j�w (`CreateObject("ADODB.Connection")`) �s�� Access ��Ʈw�A�קK�j��K�[�ѦҡC
    *   `Connect` ��k�w��{**�۰ʳЫظ�Ʈw**�\��G�p�G `DatabasePath` ���w�� `.accdb` �ɮפ��s�b�A�|���ըϥ� ADOX (����j�w) �Ыؤ@�ӷs���Ÿ�Ʈw�C
    *   `GetTableNames` ��k�ϥ� `OpenSchema` ����ϥΪ̸�ƪ�C��A�èϥ� VBA ���ت� `Collection` �ӳB�z�C��C

6.  **�A�ȼh (`PreviewService.cls`, `ImportService.cls` ��):**
    *   `PreviewService.cls` �]�t `ShowPreview` (�N Access ��ƪ�w���� Excel) �M `GetAccessTableNames` (�q `AccessDAL` �����ƪ�C��) ���޿�C
    *   `ImportService.cls` �]�t `ImportToAccess` (�B�z CSV �פJ�� Access ���޿�A�]�A�ե� `AccessDAL.DropTableIfExists` �M `AccessDAL.ExecuteSQL`)�C
    *   ��L�A�� (GL/TB/Mapping/Filter) �t�d�U�ۻ�쪺�~���޿�C

7.  **`GetTableNames` �P `ListTable` ����:**
    *   `vMain.ListTable_Enter()` Ĳ�o `RaiseEvent GetTableNames`�C
    *   `cApplication.vMain_GetTableNames()` ����ƥ�ýե� `cApplication.PopulateTableList()`�C
    *   `cApplication.PopulateTableList()` �ե� `PreviewService.GetAccessTableNames()`�A����C���M�� `vMain.ListTable`�A�T�α���A��R�s�C��A�̫᭫�s�ҥα���C

8.  **`vMain.ListTable` ComboBox �t�m:**
    *   `Style` �ݩʳ]�w�� `2 - fmStyleDropDownList`�A�ϥΪ̥u���ܡA�����J�C
    *   �w�����{���X���� `MatchRequired` �ݩʪ��]�m�C

9.  **��e���A�P����B�J:**
    *   `cApplication` �����c�w�򥻧����A�ŦX���¾�d�C
    *   ��Ʈw�۰ʳЫإ\��w��{�C
    *   `vMain.ListTable` �����ʰ��D�w�ѨM�C
    *   �w�s�W `FilterCriteria.cls` �M `FilterService.cls` ���򥻵��c�A������z��\�వ�ǳơC
    *   **�U�@�B**�O��{���M�g (Field Mapping) �\��A�A�� `vMapping.frm`, `cMapping.cls`, �M `MappingService.cls`�C

10. **�}�o�P�������:**
    *   �s�x�ϥ�**����j�w** (`CreateObject`) �H�����ۮe�ʡC
    *   �ϥ� `Debug.Print` �b VBA �ߧY��������X�ոիH���M���A�C
    *   �ϥ� `On Error GoTo Label` �i����~�B�z�A�ɦV��b���C�h�]DAL, Service�^�O���Բӿ��~�A�b�����h�]Controller�^�V�ϥΪ���ܳq�ο��~�T���C
    *   �Ҳճ����ϥ� `Option Explicit` �j���ܼ��n���C
    *   ���O�Ҳըϥ� `Private Const MODULE_NAME` �i����ѡA��K�ոտ�X�C
    *   ���O/��椺���I�s�ۨ������ɡA��ĳ�ϥ� `Me.` �����iŪ�ʡC
    *   ��`�@�P���ƥ�R�W�]�p `DoAction`, `GetInformation`�^�M�B�z�{�ǩR�W (`ObjectName_EventName`)�C
    *   ���ҲթM���n�{�ǽs�g�����ѡC

**��e����:**
��{���M�g�\��A���ϥΪ̯���N�פJ�� `GL` �M `TB` ��ƪ���������t�ιw�w�q���з����C

**�����E�J�ؼ�:**
*   ���� `vMapping.frm` �������]�p�A�T�O�]�t�Ҧ����n�� ComboBox ����ӹ��� `GLEntity` �M `TBEntity` �����C
*   �b `vMapping.UserForm_Initialize` ���A��{�q `GL` �M `TB` ��ƪ�]�γz�L Service �h�^������W�١A�ö�R������� ComboBox �U�Կ�椤�C
*   ��{ `cMapping.cls` �����ƥ�B�z�޿�A�S�O�O `vMapping_DoConfirm`�A�Ω󦬶��ϥΪ̪��M�g��ܡC
*   �}�o `MappingService.cls` �����֤��޿�A�Ω��x�s�M���γo�ǬM�g���Y�]��l���q�i��u�O�N�M�g���G�Ȧs�ζǻ��� Controller�^�C

**����n�D�P����:**

�b���R `#codebase` ���P�E�J�ؼЬ������{�� VBA �{���X��A�д��Ѩ��骺���c/�ק�/�s�W��ĳ�A��**�Y���u**�H�U�W�h�G

1.  **�{���X�ק�d��:**
    *   **�ȯ�**�ק�ηs�W `Option Explicit` ����r**����**���{���X�C
    *   **����T��**�ק�B�R���ή榡�� `Option Explicit` **���e**�����󤺮e�]�]�A `VERSION` ��B`BEGIN/END` ���B`Attribute VB_...` �浥�^�C�o�ǬO VBA ���Һ޲z�Ҳ��ݩʪ����n�����C

2.  **��X���e:**
    *   **�Ф�**�b�^�������ƶK�X���㪺�{���X�ɮשΤj�q���ק諸�{���X�C
    *   �ȴ��ѻݭn**�ק�**��**�s�W**��**����{���X���q**�C

3.  **��������:**
    *   �M������**����**�ݭn�i��o�ǭק�]�Ҧp�G�p��ŦX�s���[�c�]�p�H�p���{���`�I�����H�p�󴣰��i���@�ʡH�^�C
    *   �����ק�᪺�{���X���q**�p��B�@**�C

4.  **��I�B�J:**
    *   ����**����B�����N�Z**�������A���ɨϥΪ̦p��b Excel VBA �s�边�����Ϋ�ĳ�]�Ҧp�G�u1. �}�� `vMapping.frm`�C 2. �b `UserForm_Initialize` �ƥ󤤲K�[�H�U�{���X... 3. �}�� `cMapping.cls`...�v�^�C

5.  **��׿��:**
    *   �Y���Y�ӭ��c�I�s�b�h�إi�檺��{�覡�A�б��˧A�{��**�̨Ϊ����**�A��²�n�����A**��ܸӤ�ת��z��**�]�Ҧp�G���Ĳv�B�iŪ�ʡB�i�X�i�ʩ� VBA ������^�C

6.  **���ѭn�D:**
    *   �Ҧ��s�W�έק諸�{���X**����**�]�t**�c�餤��**���ѡC
    *   ���ѭ���ݫO���@�P�A��`�H�U�зǡG
        *   **�Ҳյ���:** �ϥ� `'=================` ���j�u�]��A�����Ҳեγ~�C
        *   **�椺����:** �ϥ� `'` ��������ν������{���X��C
        *   �������M���B²��A�����N�ϦӫD���Ƶ{���X�C