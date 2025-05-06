# JET VBA �M��

## �M�׷��[

���M�׬O�@�Өϥ� VBA (Visual Basic for Applications) �}�o�� Excel ���ε{���A���b��U�B�z�M���R�]�ȸ�ơA�S�O�O�`�b (GL) �M�պ�� (TB) ��ơC�����ѤF�@�Ӧh�B�J���ϥΪ̤����A�޾ɨϥΪ̧�����ƶפJ�B�]�w�B���ҩM���R���y�{�C��ݸ���x�s�ϥ� Microsoft Access ��Ʈw�C

## �D�n�\��

*   **CSV ��ƶפJ:**
    *   �䴩�פJ GL �M TB �� CSV �ɮסC
    *   �۰ʰ��� CSV �ɮ׽s�X (UTF-8, Big5 ��)�C
    *   �N��ƶפJ�ܫ��w�� Access ��Ʈw (`default.accdb`) ����������ƪ�C
*   **��ƹw��:**
    *   �b Excel �u�@���w���q Access ��Ʈw���J����ƪ��e (GL �� TB)�C
    *   �i�]�w�w�����̤j�C�ơC
*   **�������]�w:**
    *   ���ѨϥΪ̤��� (`vTBConfig`, `vGLConfig`) �]�w�ӷ� CSV �ɮ����P�ؼи�Ʈw��쪺�������Y�C
    *   �x�s�M�޲z�o�ǹ������Y (`MappingService.cls`)�C
*   **�������:**
    *   ���������ҵ{�ǡA�Ҧp����ʴ��� (`ValidationService.TestCompleteness`)�A��� GL �M TB ��ơC
*   **�h�B�J�ϥΪ̤���:**
    *   �z�L�D��� `vMain` �޾ɨϥΪ̧����U���ާ@�B�J�C
    *   �]�t�M�׳]�w�BTB/GL �]�w�B������ҡB�z�����]�w�����q�C

## �֤ߤ���

### �D�n���O�Ҳ� (Class Modules)

*   **`cApplication.cls`**: ���ε{�����D�n����A�t�d��զU�ӪA�ȩM�ϥΪ̤������������ʡC
*   **`AccessDAL.cls`**: ��Ʀs���h (Data Access Layer)�A�ʸˤF�Ҧ��P Access ��Ʈw�������޿� (�s�u�B���� SQL�BŪ����Ƶ�)�C
*   **`ImportService.cls`**: �B�z CSV �ɮ׶פJ�� Access ��Ʈw���޿�C
*   **`PreviewService.cls`**: �t�d�q Access ��ƮwŪ����ƨæb Excel �u�@����ܹw���C
*   **`MappingService.cls`**: �޲z�M�x�s GL �� TB �����������Y�C
*   **`ValidationService.cls`**: �]�t������Ҫ������޿�A�Ҧp����ʴ��աC
*   **`GLService.cls` / `TBService.cls`**: ���O�B�z GL �M TB �S�w���~���޿� (�ثe������¦)�C
*   **`GLEntity.cls` / `TBEntity.cls`**: �w�q GL �M TB ��ƪ����鵲�c�C
*   **`AppConfig.cls`**: (����) �Ω��x�s���ε{�����]�w�M�ѼơC

### �D�n���Ҳ� (Form Modules)

*   **`vMain.frm`**: ���ε{�����D�����A���ѦU�D�n�\�઺�J�f�C
*   **`vProject.frm`**: �M�׳]�w���������C
*   **`vTBConfig.frm`**: TB ��ƶפJ�M�������]�w���C
*   **`vGLConfig.frm`**: GL ��ƶפJ�M�������]�w���C
*   **`vValidation.frm`**: ������Ҭ����ާ@�����C
*   **`vCriteria.frm`**: (����) �Ω�]�w�z����󪺪��C

### �зǼҲ� (Standard Modules)

*   **`mod_Utility.bas`**: �]�t�q�Ϊ����U��ơA�Ҧp `Start` �{�� (�Ұ����ε{��) �M `DetectCSVEncoding` (���� CSV �s�X)�C

## �u�@�y�{ (Workflow)

���ε{�����嫬�u�@�y�{�j�P�p�U (�� `cApplication.cls` ����)�G

1.  **�Ұ����ε{��**: �z�L���� `mod_Utility.Start` �{�ǨӪ�l�ƨ���ܥD���� `vMain`�C
2.  **�B�J 1: �M�׻P��ƶפJ�]�w**
    *   �ϥΪ̳z�L `vMain` �i�J�B�J 1�C
    *   **�M�׳]�w (`vProject`)**: (����\��ݽT�{)
    *   **TB �]�w (`vTBConfig`)**:
        *   �פJ TB CSV �ɮסC
        *   �w���פJ�� TB ��ơC
        *   �]�w TB �������C
    *   **GL �]�w (`vGLConfig`)**:
        *   �פJ GL CSV �ɮסC
        *   �w���פJ�� GL ��ơC
        *   �]�w GL �������C
3.  **�B�J 2: ������� (`vValidation`)**
    *   ����U�ظ�����Ҵ��աA�Ҧp�G
        *   ����ʴ��աC
        *   ��󥭿Ŵ��աC
        *   RDE ���աC
        *   ��ع����C
4.  **�B�J 3: �z����� (`vCriteria`)**
    *   �]�w�Ω������R�γ����ͪ��z�����C
5.  **�B�J 4: (�ݩw�q)**
    *   ���򪺸�ƳB�z�Τ��R�B�J�C

## �p��ϥ�

1.  �}�� `JET.xlsm` �ɮסC
2.  (�p�G�ݭn) �ҥΥ����C
3.  �w���|���@�ӫ��s�Τ覡��Ĳ�o `mod_Utility.Start` �{�ǥH�Ұ����ε{���D�����C

## �`�N�ƶ�

*   ���M�ר̿� Microsoft Access Database Engine�C�нT�O�w�w�ˬ��������� Access Database Engine (�Ҧp Microsoft.ACE.OLEDB.12.0)�C
*   �����\�� (�p `PreviewService` ���]�w�u�@�� CodeName) �i��ݭn�ҥ� "�H�� VBA �M�ת���ҫ��s��" (�b Excel �ﶵ -> �H������ -> �H�����߳]�w -> �����]�w��)�C
*   ��Ʈw�ɮ� `default.accdb` �w���P `JET.xlsm` �ɮצ��P�@�ؿ��U�C

---
*�� README.md �ɮ׬O�ھڵ{���X�w�۰ʲ��ͪ���B�����A�i��ݭn�i�@�B����ʽվ�M�ɥR�C*