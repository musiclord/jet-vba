# ���O�y�z
- **mod_Utility.bas**
    - �q�Υ\��u��A�@���}�������Ҧ����O�s���ϥΡC

- **vMain.frm**
    - �D�{�������A���ϥΪ̰��� :1.�פJ�ɮ�,2.���Ҹ��,3.�z�����,4.��X���i�F�åB�]�p�w���\��A��ܨé�u�@���˵���ƪ��e1000����ơC

- **vMapping.frm**
    - �M�g��줶���A���ϥΪ̱N�פJ����ơA�ǥѾާ@�U�Ԧ����ӬM�g�ܥ��T���������W�١C

- **cApplication.cls**
    - ��VBA���ε{�Ǫ��D�n����A���� `vMain` �������{�ǡA�éI�s������ơC

- **cMapping.cls**
    - �M�g��쪺����A���� `vMapping` �������{�ǡC

- **AccessDAL.cls**
    - �t�d�Ҧ��P Microsoft Access ��Ʈw�����ʡA�ʸ� ADO/ADOX �s�u�A���� SQL �y�y ( *�p SELECT, INSERT, UPDATE, DELETE* ) �H�ξާ@��Ʈw���� ( *�p�ˬd���O�_�s�b�B�إߪ��B�s�W��쵥* ) �����h�Ӹ`�C

- **GLEntity.cls**
    - `GL` (General Ledger) ��ƹ������O�A�w�q `GL` ��Ƶ��c�����ҡC

- **TBEntity.cls**
    - `TB` (Trial Balance) ��ƹ������O�A�w�q `TB` ��Ƶ��c�����ҡC

- **ImportService.cls**
    - �B�z�ɮ׶פJ���A�ȼh�A�t�d�q�����פJ����ɮ�(�pCSV�BExcel)��Access��Ʈw�A�PAccessDAL��@������Ʀs���ާ@�C

- **MappingService.cls**
    - �B�z���M�g (Field Mapping) �޿誺�A�ȼh�A�Ω�зǤ����W�١A�x�s�P�޲z�M�g���Y�C

- **PreviewService.cls**
    - �B�z�w����ƪ��A�ȼh�A�t�d�N��Ʈw�����w��ƪ��e1000����ơA���J�ܫ��w���u�@��C

# �y�{�y�z
- �פJ�ɮ�
    - �פJ `GL.csv` �� Access ����ƪ� (table) `GL`
    - �פJ `TB.csv` �� Access ����ƪ� `TB`
    - �w����ƪ� `GL` ��u�@�� (worksheet) `GL`
    - �w����ƪ� `TB` ��u�@�� `TB`
    - �ާ@�U�Ԧ����t����� (Field Mapping)
    - �W�[��� [��󶵦�] ����ƼW�j
    - �зǤƸ�ƪ� `GL` �M `TB` �� `GL#` �M `TB#`
- ���Ҹ��
    - ��������� (Completeness)
    - �ɶU�������� (Document Not Balance)
    - ��Ƥ����������� (Relevant Data Elements)
    - ��ذt�� (Account Mapping)
- �z�����
    - �]�w�򥻿z����� (Criteria Selection)
    - �զX�z�����]���@�տz��t�m
    - �̷ӿz��t�m�i��z��@�~
    - �w���z�ﵲ�G��u�@��
- ��X���i
    - �N���ҵ��G `#Completeness` ��ܽd���u�@�� `Validation.xlsx`
    - �N�z�ﵲ�G `#Filtered` ��ܽd���u�@�� `Filtered_Result.xlsx`

# �{�Ǵy�z
	1.	�פJGL.csv��Access��Ʈw
		1.1	�ˬd�ڥؿ��O�_�s�b�P�W����Ʈw
		1.2	�ˬd��Ʈw�O�_�s�b�P�W����ƪ�
		1.3	�H���T���s�X�θ�ƫ��O�פJ��ƪ�
	2. Ū��Access��Ʈw�A�N�s�פJ(�Τw�s�b�P�W)����ƪ���J�ܤu�@��
		2.1	�ˬd�O�_�s�b�P�W���u�@��
		2.2	�мg�¸�ơA�פJ�s���
		2.3	�N�u�@��@���d�ߵ��G��View�A�H1000����Ƭ���
	3. �ϥΪ̮ھ�vMapping���U�Ԧ����A�N mGLEntity �w���w�q�����W�١A�t��ܶפJ��GL��ƪ�
		3.1 �b mGLEntity ���O����ƪ����W��
		3.2 �N�ϥΪ̭��s�t�諸���O���b mGLEntity �ós�������W�١A�ϱo����ETL�i�H���T���ެM�g�᪺���
		3.3 c/ c
	4. ����ƼW�j�A�NGL��ƪ�K�["LineItem"���A�ΨӰϧO���ƬۦP�ǲ����X�����P����
		4.1 �I�s mGLEntity �ӥH�w�q�n����Ƶ��c�A�I�s cAccess �Ӿާ@���骺�d�߻y�y
		4.2 �ϥ�SQL�s�W��� "LineItem" ��C�ӬۦP "�ǲ����X" ��������ƶi��v�@�W�[�����Ǹ�
		4.3 ��Ʈw��GL��ƪ��s�F
	5. �˵���s�᪺��ƪ�
		5.1 �ϥΪ��I�� vMain �� ListTable ���ӿ���Q�w������ƪ�
		5.2 �ϥΪ��I�� vMain �� ButtonPreview ��Ĳ�o DoPreview�ƥ�
		5.3 �ЫةΧ�s(�Y�w�s�b)�u�@�� "Preview"
		5.4 �N�n�w����ƪ��d�ߵ��G���J�ܤu�@��


�b vMain.frm �������s ButtonPreview �b�I����|Ĳ�o DoPreview �ƥ�å� cApplication.cls �� vMain_DoPreview �B�z�Өƥ�A�b�y�{�W�O�b�d�߸�Ʈw�öǦ^�ܤu�@��A�]���L���ӯ୫�ƨϥΡA�N mImport.LoadToExcel ��g�A�ϥλP vMain_DoPreview �ۦP���B�@�޿�C�аݭY�n�s�W�@�Ө�ƨӳB�z�W�z�\��A�L�����k�ݩ�������O? �p��ե�?

���F�����W�z���e�A�Х����ڳ]�p #file:cAccess.cls  �çi�D�ڨ�L�ϥθ�Ʈw����ƻP���O�n�p��ե�? �t���������ǳ����ݭn���? �Ҧp #file:mImport.cls  ����k�B�άO #file:mGLService.cls  �|�������� AddLineItem() �޿�B�H��MVC�[�c���� #file:mGLEntity.cls  ���C�b��Ҫ��L�{���A�ХH�A�{�����u�ƥB²�䪺�N�X���D�A�קK�L�׳]�p�������޿�A���ɥN�X��Ū�ʡA�åB�A�ݭn�H��e�M�ת������Ӥ��R�N�X����O�_�ŦXExcel��VBA�C