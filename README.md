# jet-vba

���M�צ��b�ϥ� Excel VBA �}�o��O�b�������� (Journal Entry Testing, JET) �u��C

## �[�c

���M�׿�`�� MVC ��h�ҵo�����h�[�c�A�H�P�i���`�I�����M�i���@�ʡG

*   **���� (`vMain.frm`, `vMapping.frm`):** �Ω󤬰ʪ��ϥΪ̤������C
*   **��� (`cApplication.cls`, `cMapping.cls`):** �B�z UI �ƥ�A������ε{���y�{�A�ñN���ȩe�����A�ȼh�C
*   **�A�ȼh (`ImportService.cls`, `PreviewService.cls`, `GLService.cls`, `TBService.cls`, `MappingService.cls`):** �ʸ˯S�w�\�઺�~���޿�]�פJ�B�w���B��ƳB�z�B�M�g�^�C
*   **��Ʀs���h (DAL) (`AccessDAL.cls`):** �޲z�P Microsoft Access ��Ʈw (`.accdb`) ���Ҧ����ʡA�ϥ� ADODB �M ADOX �ñĥΫ���j�w�C
*   **���ε{�� (`mod_Utility.bas`):** �]�t�q�λ��U��ơ]�Ҧp�G�ɮ׿�ܡBCSV �s�X�����^�C

## �D�n�\��

*   **CSV �פJ:** �N�`�b (GL) �M�պ�� (TB) ��Ʊq CSV �ɮ׶פJ Access ��Ʈw (`default.accdb`)�C
    *   �B�z���P�� CSV �s�X�]���� UTF-8 BOM�A�Y�������ѫh�w�]�� 950�^�C
    *   �פJ�ɦ۰ʧR���í��s�إߥؼи�ƪ� (`GL`, `TB`)�C
*   **��Ʈw�޲z:**
    *   �������ճs���ɡA�p�G `poc` �ؿ������s�b `default.accdb` Access ��Ʈw�ɮסA�h�۰ʫإ߸��ɮסC
*   **��ƹw��:** �N��w Access ��ƪ���Ƹ��J�M�Ϊ� Excel �u�@��]�Ҧp�G`GL_Preview`, `TB_Preview`�^�ѨϥΪ��˵��C
    *   �N�w������i�]�w���C�� (`MAX_ROWS_TO_SHOW`)�C
*   **�ʺA��ƪ�M��:** �I���U�ԫ��s�ɡA�ϥ� Access ��Ʈw���i�Ϊ��ϥΪ̸�ƪ��R `vMain` �W�� ComboBox (`ListTable`)�C
*   **���M�g:** �Ω�N�ӷ����M�g��ؼ���쪺��¦�[�c�]��@�Ӹ`�b `cMapping`, `vMapping`, `MappingService` ���^�C
*   **��ƳB�z:** �Ω�B�z�פJ��ƪ��A�� (`GLService`, `TBService`)�]�Ӹ`�ݶi�@�B�w�q/��@�^�C

## �ثe���A (�I�� 2025�~4��24��)

*   �֤߬[�c�]���ϡB����B�A�ȡBDAL�^�w�إߡC
*   CSV �פJ�B�۰ʸ�Ʈw�إߩM�򥻸�ƹw���\��w��@�C
*   `cApplication.cls` �����c�w�j�P�����A�H�ŦX���¾�d�C
*   **�ثe�J�I/���D:** �ѨM `vMain` �W�� `ListTable` ComboBox �b�I���U�ԫ��s�ɵL�k��ܨ�U�ԦC�����D�A���ީ��h���ؤw�z�L `GetTableNames` �{�ǥ��T��s�C

## �]�w�P�ϥ�

1.  �b Microsoft Excel ���}�� `poc/JET.xlsm` �ɮסC
2.  �p�G�X�{���ܡA�бҥΥ����C
3.  �D���� (`vMain`) ���ӷ|�X�{�C
4.  �ϥΫ��s�פJ GL/TB CSV �ɮס]�T�O�d�� CSV ��� `poc/data` ��Ƨ����^�C
5.  �I����ƪ�M�� ComboBox �W���U�Խb�Y�H��s���˵��i�Ϊ���ƪ�]�ثe�J��U����ܰ��D�^�C
6.  ��ܤ@�Ӹ�ƪ�]�p�G�i��^���I���u�w���v�H�b Excel ���˵���ơC
7.  �p�G `poc` ��Ƨ������s�b `default.accdb` ��Ʈw�A���N�Q�۰ʫإߡC

## �޳N

*   Microsoft Excel VBA
*   Microsoft ActiveX Data Objects (ADODB) - ����j�w
*   Microsoft ADO Ext. for DDL and Security (ADOX) - ����j�w
*   Microsoft Scripting Runtime (FileSystemObject) - ����j�w