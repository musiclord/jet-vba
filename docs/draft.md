purpose, objectvie: Solution for what? �����Ƥ��u ���ɮįq 
better solution(�u��{�p)? Time cost, user experience?
compete? KCT? or else, target aimed for "2025 JET �Ҥ��ؼ�"
idea limitation? �۬� core target for?
application -> annoymys or virtual case fro demonstration.
Target to design a prototype, extend to visions. blah blah blah.



�ھ� 0321 & 0328 �y�z�����ҧ����H�U�ݨD
- Import Data (CSV, TXT)
	- Import to worksheet from PowerQuery
	- Worksheet format
- Column Selection
	- Dropdown List for columns
	- 
- Account Mapping (track to histroy, need state model)
	- Difference, records,
- Validate data Completeness
	- 
- Criteria Matching (5-7 limit fields)
	- Limiation resources, simplest logics ( such as dates, account filtering)
- Export WP (Most simplified format)
	- Worksheet content format, 
	
	
�T�{ JET ����y�{
JET ��ƶq �d��? �榡?

�ڥ��b�B�zJournal Entry Testing Tool(JET)���ץ�A�n�N������u��E���ܷs���ѨM��סA���ثe���b�������q�F�������׬O�ϥ�Caseware Idea�æb�����W�}�oVBA�@���۩w�q�u��A�t�X�ϥ�²����Windows Forms�@�������C�{�b�ѩ���v�L���ǳƭn�^�OIdea�A�]���b�����p��b���̿�Idea�����ҤU�}�oJET�A�ثe����Ӥ�V:
1. �̿�� .NET �� Windows Forms Application
2. �̿�� Office���ε{���� VBA ����
�ѩ�ثe�귽�j�h���ೡ�ݽ����B�L�h�̿઺���ҡA�]���|���V��VBA�����A���O�����󩳼h�䴩�AExcel���䴩�ާ@�W�L�@�ʸU����ƪ����e�A�Ӧ��ɭ�general ledger�|�W�L�o�ӼƦr�A�Ӥ@��JET�b�����Ʊ��O�H�U:
1. �פJ�ɮ�(.csv, .txt, .xlsx)
2. ��ƬM�g�A�]���Ȥᴣ�Ѫ����(�Ҧpgeneral ledger��trial balance��)�|�ھڤ��P�t�λP���q�Ӧ����P���R�W�A�Ҧp"�ǲ��s��"�i��b������R�W��"DocumentNo."�Ψ�L�W�١A���F�T�O����ETL���L�{���T�A�|�ݭn��w�@�ӼзǤƪ���ơA�]�N�O���g�L���M�g(column mapping)�Ψ�L�A�{����A�X���W�٨Ӵy�z�o�Ӭy�{�C
3. 


pbi�C�ӨB�J�|�O����}��
�Ҧp�bpower query�����ާ@�B�J�A�|��idea���ˬ���
���t����D�A���GA�B�z�A�]�����ݦҼ{�۩w�q���

office script(typescript)�i�H�B�z�ާ@�B�J�����P�٭��?
�B�zJET�ɦ�����RPA�y�{�A�Ҧp�ɤJ��Ʀp�G�ۦP�A�O�_�i�HRPA�e�m�@�~?


�s�@definition��A�����C�ӼзǤƦW�٪��w�q�A�Ҧp:
�ǲ��s��:���ѩM�l?�C�@���|�p������ߤ@�ѧO�X

VBA �O VB ���l���A�O�J�b Microsoft Office ���ε{���]�p Excel�BWord�BAccess�^���A�Ω�۰ʤƳo�����ε{�������ȡC


C#��B�z�j�q��ƶ�?
�p���x�s��ƪ�? �p��w�qERM?


ASP.NET Core + WebAssembly + Blazor + Tailwind CSS

��e��s�x�j�ҵ{ JE Tool ������ ���ӬO���� timesheet 


JET ��� �i�঳ ��ƫ��O ����줣�@�� �Ҧp�ɶU���t�� �H���P�Φ����
���ݭn�u�ƪ���

���ӨB�J �n�x�s���A ����B�J���_ �o���s�}�l


�B�J�@ ���Cimport ���e �Ҧp���t��
�B�J�G mapping��� �N�P�ʽ���� �@���Pgroup�����O �]���z��� �|�H�P�ʽ����O ��key
�B�J�T ���z �t�z(except, exclude) ������ �Ҧp ��X�P������ñư����w���


aware of enegagement number, �N�U�O�M�׿W�ߥX���ҡA�Ҧplocal DB�A�ݦU�O�]�p��ƪ� (�קK��Ʈ���)

datagridview�i�H�w����ƪ��e


criteria: 

����� ������b���W �b���W����o���W

2025-3-26
je testing optimiz
sop select by each year, for exmaple 5 criteria, then define as preset has those 5 criteria
customize criteria already in current version jet

date keyin by use input (which method) --> Transform to valid adta format

category is it neeceesary? --> optional column, 

isManual --> excel predifined --> load in for Step1_Check_User_Define_Manual as one column --> 

Criteria --> contains (logic needs to be re-defined) --> 

DropDownList -->Catch Event by data(new item) --> Handle event if catch --> 

question 12 useinput by string --> avoid string length out f forms --> let worksheet as default view and extract as data input.

Document, LineItem--> IF-ELSE to check if data containes necessary  --> for example if LineItem not in GL, then created LineItem group by DocumentNumber

Draw the complete flow chart of "validation" procedure. 

�毸���ǲ�?

AND/OR criteria --> Re-design a QUeryBuilder that can modify not only ONE but MANY criteria at ONCE.

�n�ާ@���D��: GL, TB, 
�n�ާ@�����: ���, ������, �ײv, 
�n�ާ@���޿�: <= , >= , && , + , - , * , / ,