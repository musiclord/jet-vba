# �ݶ}�o���{��:
- �פJ�ɮ�
    - �T�O GL �M TB ����ƪ��T�פJ�A�æb Access ��Ʈw���s����������ƪ�
- ���Ҹ��
    - ����ʴ���
        - �w�q����ʴ��ժ��޿�M�y�{
        - �C�X�Ӵ��ժ��ݨD�M���
        - ��{����ʴ���
    - �ɶU����
        - �w�q�ɶU�������ժ��޿�M�y�{
        - �C�X�Ӵ��ժ��ݨD�M���
        - ��{�ɶU��������
    - ��ذt��
        - �w�q��ذt�諸�޿�M�y�{
        - �C�X�ӥ\�઺�ݨD�M���
        - ��{��ذt�諸�\��
        - ��{��ذt�諸����
- �z�����
    - ��{

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
    - ���ƪ� `GL` �M `TB` �H�U�Ԧ����t����� (Field Mapping)
    - ���ƪ� `GL` �W�[��� [��󶵦�] ����ƼW�j
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


