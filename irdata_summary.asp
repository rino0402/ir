<%
Option Explicit
Response.Buffer = false		'�y�[�W�o�͂��o�b�t�@�Ɋi�[���邩�ǂ���
Response.Expires = -1		'�u���E�U��ɃL���b�V�������y�[�W�̗L���������؂��܂ł̎���
%>
<%
Function GetVersion()
	GetVersion = "2012.05.04 �V�K�쐬"
	GetVersion = "2012.05.07 �u���ʂ��R�s�[�v�Ō������ʂ��N���b�v�{�[�h�ɃR�s�[����悤�ɑΉ�"
	GetVersion = "2012.05.07 ���Ƌ敪�� �v�捷(����/���v)�̌v�Z�����C��"
	GetVersion = "2012.05.07 ���Ƌ敪�T���� �̑Ή�"
	GetVersion = "2012.05.08 �݌v�̌����͈͂�201104�`�ɂȂ��Ă����̂�201204�`�ɏC��"
	GetVersion = "2012.05.11 �o�c�T�� ���Ƌ敪�ʂɁu_���̑��v��ǉ�"
	GetVersion = "2012.12.15 �o�c�T�� 2013�N�x�ł̏W�v�Ή�"
	GetVersion = "2012.12.20 �C���^�[�t�F�[�X���P��...�O���ϐ��폜"
	GetVersion = "2012.12.20 �o�͌`��:���Ƌ敪�T����(�N��-����) �ǉ�"
	GetVersion = "2013.01.30 �o�͌`��:���Ƌ敪�T����(���n) �v��1�`3���̑Ή�"
	GetVersion = "2013.05.10 �o�͌`��:���Ƌ敪�T����(�Ԑڂ��莞�ԃ`�F�b�N) ���e�펞�Ԃ̃`�F�b�N�Ɏg�p���ĉ�����"
	GetVersion = "2013.05.11 �o�͌`��:�ꗗ�\ �Ȗ�/���Ƌ敪���Ƃɒl���`�F�b�N�ł��܂��B"
	GetVersion = "2013.05.28 ���Ƌ敪=�� �Ō��������ꍇ�A�e���Ԃ̒l���傫���Ȃ�s����C��"
	GetVersion = "2013.07.23 ���x�ł̏W�v�^�����ɑΉ�"
	GetVersion = "2013.11.15 �o�͌`��:���Ƌ敪�T����(���n:�v��11�`3��)�̑Ή�"
	GetVersion = "2014.07.25 �o�͌`��:���Ƌ敪�T����(���n:�v��(7,8,9,10�`3))�̑Ή�"
	GetVersion = "2014.11.28 �o�͌`��:���Ƌ敪�T����(�N��) ���̕\���ԈႢ�����"
	GetVersion = "2015.01.20 �o�͌`��:�N�Ԏ��Ƌ敪�ʊT����(����) ���ԃ`�F�b�N����ǉ��i��ƒ��E�E�E)"
	GetVersion = "2015.01.23 �o�͌`��:���Ƌ敪�T����(���n:�v��)�̕s��C��"
	GetVersion = "2015.09.17 �o�O�F�N�Ԏ��Ƌ敪�ʊT���� �� �O�N�v��ƍ����v��ɂȂ��Ă���"
	GetVersion = "2015.09.18 ���Ƌ敪�T����(�N��)�F�������ʂ����O�N���тɂȂ��Ă����s��C��"
	GetVersion = "2017.10.30 �o�c�T��:���㌴���^�e���v�^�o������/���E���v�^�Œ��ɕύX"
	GetVersion = "2017.11.14 �C����(80% �����ɑ���m(__)m)...�N�Ԏ��Ƌ敪�ʊT����(����) �������ʂ�/�����v��"
	GetVersion = "2017.11.15 �����v�悪�������W�v�ł��Ȃ��s����C��"
	GetVersion = "2017.11.15 �����v��u���̑��v�ŃG���[����������s����C��"
	GetVersion = "2018.11.13 �N�Ԏ��Ƌ敪�ʊT����(����)�ŁuNull �l�̎g�������s���ł��v����������s��C��"
	GetVersion = "2020.01.27"
End Function

Function GetDbName()
	GetDbName	= "IR"
End Function

%>
<%
Function YKSub(byVal strCenterCD,byVal strYM,byVal strYM2,byval strJKubun,byval strKubun)
	dim	s
	s = "select"
	s = s & " sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst04"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst05"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst06"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst07"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst08"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst09"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst10"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst11"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst12"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst01"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst02"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst03"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) APln04"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) APln05"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) APln06"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) APln07"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) APln08"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) APln09"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) APln10"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) APln11"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) APln12"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) APln01"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) APln02"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) APln03"

	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst04"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst05"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst06"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst07"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst08"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst09"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst10"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst11"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst12"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst01"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst02"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst03"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) BPln04"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) BPln05"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) BPln06"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) BPln07"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) BPln08"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) BPln09"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) BPln10"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) BPln11"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) BPln12"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) BPln01"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) BPln02"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) BPln03"

	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst04"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst05"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst06"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst07"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst08"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst09"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst10"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst11"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst12"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst01"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst02"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst03"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C1Pln04"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C1Pln05"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C1Pln06"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C1Pln07"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C1Pln08"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C1Pln09"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C1Pln10"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C1Pln11"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C1Pln12"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C1Pln01"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C1Pln02"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C1Pln03"

	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst04"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst05"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst06"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst07"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst08"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst09"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst10"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst11"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst12"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst01"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst02"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst03"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C2Pln04"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C2Pln05"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C2Pln06"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C2Pln07"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C2Pln08"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C2Pln09"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C2Pln10"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C2Pln11"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C2Pln12"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C2Pln01"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C2Pln02"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C2Pln03"

	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst04"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst05"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst06"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst07"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst08"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst09"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst10"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst11"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst12"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst01"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst02"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst03"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C9Pln04"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C9Pln05"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C9Pln06"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C9Pln07"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C9Pln08"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C9Pln09"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C9Pln10"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C9Pln11"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C9Pln12"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C9Pln01"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C9Pln02"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C9Pln03"

	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst04"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst05"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst06"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst07"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst08"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst09"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst10"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst11"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst12"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst01"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst02"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst03"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) DPln04"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) DPln05"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) DPln06"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) DPln07"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) DPln08"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) DPln09"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) DPln10"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) DPln11"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) DPln12"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) DPln01"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) DPln02"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) DPln03"

	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst04"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst05"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst06"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst07"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst08"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst09"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst10"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst11"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst12"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst01"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst02"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst03"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) X2Pln04"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) X2Pln05"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) X2Pln06"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) X2Pln07"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) X2Pln08"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) X2Pln09"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) X2Pln10"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) X2Pln11"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) X2Pln12"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) X2Pln01"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) X2Pln02"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) X2Pln03"

	s = s & " from IrData"
	s = s & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM2,3) &"'"
	s = s & "   AND CenterCD = '" & strCenterCD &"'"
	if strJKubun <> "" then
		s = s & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
	else
		s = s & vbcrlf & "   AND SyushiCd <> ''"
	end if
	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if
	s = s & SqlWhere("and", "SyushiCd", strSyushiCD)
'	if strSyushiCD <> "" then
'		s = s & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'	end if
	YKSub = s
End Function
Function TMSub(byVal strCenterCD,byVal strYM,byVal strYM2,byval strJKubun,byval strKubun,byval strFlg)
	dim	s
	s = "select"
	s = s & " sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst04"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst05"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst06"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst07"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst08"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst09"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst10"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst11"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst12"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst01"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst02"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst03"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) APln04"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) APln05"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) APln06"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) APln07"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) APln08"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) APln09"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) APln10"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) APln11"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) APln12"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) APln01"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) APln02"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) APln03"
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst04"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst05"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst06"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst07"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst08"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst09"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst10"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst11"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst12"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst01"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst02"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst03"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM101Pln04"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM101Pln05"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM101Pln06"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM101Pln07"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM101Pln08"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM101Pln09"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM101Pln10"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM101Pln11"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM101Pln12"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM101Pln01"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM101Pln02"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM101Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst04"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst05"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst06"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst07"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst08"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst09"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst10"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst11"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst12"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst01"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst02"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst03"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM102Pln04"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM102Pln05"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM102Pln06"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM102Pln07"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM102Pln08"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM102Pln09"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM102Pln10"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM102Pln11"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM102Pln12"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM102Pln01"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM102Pln02"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM102Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst04"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst05"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst06"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst07"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst08"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst09"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst10"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst11"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst12"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst01"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst02"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst03"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM201Pln04"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM201Pln05"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM201Pln06"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM201Pln07"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM201Pln08"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM201Pln09"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM201Pln10"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM201Pln11"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM201Pln12"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM201Pln01"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM201Pln02"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM201Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst04"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst05"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst06"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst07"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst08"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst09"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst10"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst11"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst12"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst01"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst02"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst03"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM202Pln04"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM202Pln05"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM202Pln06"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM202Pln07"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM202Pln08"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM202Pln09"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM202Pln10"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM202Pln11"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM202Pln12"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM202Pln01"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM202Pln02"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM202Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst04"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst05"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst06"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst07"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst08"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst09"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst10"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst11"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst12"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst01"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst02"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst03"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM301Pln04"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM301Pln05"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM301Pln06"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM301Pln07"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM301Pln08"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM301Pln09"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM301Pln10"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM301Pln11"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM301Pln12"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM301Pln01"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM301Pln02"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM301Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst04"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst05"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst06"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst07"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst08"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst09"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst10"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst11"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst12"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst01"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst02"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst03"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM302Pln04"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM302Pln05"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM302Pln06"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM302Pln07"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM302Pln08"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM302Pln09"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM302Pln10"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM302Pln11"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM302Pln12"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM302Pln01"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM302Pln02"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM302Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst04"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst05"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst06"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst07"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst08"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst09"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst10"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst11"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst12"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst01"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst02"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst03"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM401Pln04"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM401Pln05"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM401Pln06"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM401Pln07"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM401Pln08"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM401Pln09"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM401Pln10"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM401Pln11"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM401Pln12"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM401Pln01"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM401Pln02"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM401Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst04"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst05"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst06"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst07"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst08"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst09"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst10"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst11"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst12"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst01"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst02"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst03"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM402Pln04"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM402Pln05"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM402Pln06"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM402Pln07"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM402Pln08"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM402Pln09"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM402Pln10"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM402Pln11"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM402Pln12"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM402Pln01"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM402Pln02"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM402Pln03"
end select
	s = s & " from Attendance"
	s = s & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM2,3) &"'"
	s = s & "   AND CenterCD = '" & strCenterCD &"'"
	if strJKubun <> "" then
		s = s & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
	else
		s = s & vbcrlf & "   AND SyushiCd <> ''"
	end if
	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if
	s = s & SqlWhere("and", "SyushiCD", strSyushiCD)
'	if strSyushiCD <> "" then
'		s = s & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'	end if
	TMSub = s
End Function
Function GetYm()
	dim	intYear
	dim	intMonth
	dim	intDay
	dim	dt

	dt = Now()
	intYear		= Year(dt)
	intMonth	= Month(dt)
	intDay		= Day(dt)

	if intDay < 25 then
		intMonth = intMonth - 1
		if intMonth < 1 then
			intMonth = 12
			intYear	= intYear - 1
		end if
	end if
	GetYm = intYear & right("0" & intMonth,2)
End Function
'-----------------------------------------------------------
'�o�͌`�����X�g��Ԃ�
'-----------------------------------------------------------
Function GetPTypeList()
	dim	strPTypeList
	strPTypeList = ""
	strPTypeList = strPTypeList & "<!-- GetPTypeList() start -->" & vbCrLF

	dim	strPType
	strPType = GetRequest("ptype","")

	strPTypeList = strPTypeList & GetOptionTag("pTable"			,"�o�c�T��"					,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableSyushi"	,"�o�c�T��+���x��"			,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJM"		,"���Ƌ敪�ʓ����ڍ�"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJK"		,"���Ƌ敪�T����"			,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKKan"	,"���Ƌ敪�T����(���Ԏ��ԃ`�F�b�N)"	,strPType) & vbCrLF
	strPTypeList = strPTypeList & "<OPTGROUP label=""���Ƌ敪�T����(���n)"">"
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku7"	,"���n:�v��(7�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku8"	,"���n:�v��(8�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku9"	,"���n:�v��(9�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku10"	,"���n:�v��(10�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku11"	,"���n:�v��(11�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku"	,"���n:�v��(12�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku1"	,"���n:�v��(1�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku2"	,"���n:�v��(2�`3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku3"	,"���n:�v��(3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & "</OPTGROUP>"
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYear"	,"���Ƌ敪�T����(�N��)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYearMonth2"	,"���Ƌ敪�T����(�N��-����) �������ʂ�/�����v��",strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYearMonth"	,"���Ƌ敪�T����(�N��-����) �������ʂ�/�����v��",strPType) & vbCrLF

	strPTypeList = strPTypeList & GetOptionTag("pList"			,"�ꗗ�\"					,strPType) & vbCrLF

	strPTypeList = strPTypeList & "<!-- GetPTypeList() end -->"
	GetPTypeList = strPTypeList
End Function
%>
<!--#include file="makeWhere.asp" -->
<%
'------------------------------------------------------
'��������
'------------------------------------------------------
if GetRequest("submit1","") <> "" then
	Call SetCookie("CenterCD")
end if
'------------------------------------------------------
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
<TITLE>�o�c����-�o�c�T��</TITLE>
<!-- jdMenu head�p include �J�n -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head�p include �I�� -->
</HEAD>
<SCRIPT LANGUAGE="JavaScript"><!--
	function DoCopy(arg){
		var doc = document.body.createTextRange();
		doc.moveToElementText(document.all(arg));
		doc.execCommand("copy");
		window.alert("�N���b�v�{�[�h�փR�s�[���܂����B\n�\��t���ł��܂��B" );
	}
--></SCRIPT>
<BODY>
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc-ir.asp" -->
<!-- jdMenu body�p include �I�� -->
  <FORM name="sqlForm"> <!--accept-charset="UTF-8"-->
	<table id="sqlTbl">
		<caption style="text-align:left;">�o�c����-�o�c�T��</caption>
		<tr>
			<th>�W�v�N��</th>
			<th>�Z���^�[</th>
			<th>���Ƌ敪</th>
			<th>���x</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT class="input" TYPE="text" NAME="YM" VALUE="<%=GetRequest("YM",GetYM())%>" size="8" style="text-align:center;" required pattern="^[0-9]+$"><!-- placeholder="�N��(yyyymm)�����"-->
			</td>
			<td align="center">
				<select class="input" NAME="CenterCD">
				<%=GetCenterList()%>
				</select>
			</td>
			<td align="center">
				<select class="input" NAME="JKubun">
				<%=GetJKubunList()%>
				</select>
			</td>
			<td align="center">
				<INPUT class="input" TYPE="text" NAME="SyushiCD" VALUE="<%=GetRequest("SyushiCD","")%>" style="text-align:center;" placeholder="���x ex.111,112"><!-- size="8" -->
			</td>
		</tr>
		<tr>
			<td colspan="4" nowrap>
				<label for="ptype">�@�o�͌`���F</label>
				<select class="input" NAME="ptype" id="ptype">
				<%=GetPTypeList()%>
				</select>
			</td>
		</tr>
		<tr bordercolor=White>
			<td colspan="4">
				<INPUT class="cssbutton" TYPE="submit" value="����" id=submit1 name=submit1>
				<INPUT class="cssbutton" TYPE="reset" value="���Z�b�g" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
				<span class="info_new"><%=GetVersion()%></span>
			</td>
		</tr>
	</table>
  </FORM>
<%	if len(GetRequest("submit1","")) > 0 then
		Server.ScriptTimeout = 3000
%>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<div>
		<INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			 value="������...ScriptTimeout=<%=Server.ScriptTimeout%>" id="cpTblBtn" disabled>
	</div>

	<div id='resultDiv'>
	<TABLE id="resultTbl">
		<caption  style="text-align:left;"><%=GetCaption(GetRequest("YM",GetYM()),GetRequest("CenterCd",""),GetRequest("JKubun",""),GetRequest("ptype","pTable"))%></caption>
	<%
'		Response.Flush
		dim	objDb
		Set objDb = Server.CreateObject("ADODB.Connection")
		objDb.Open GetDbName()
	%>
		<thead>
			<%=MakeHeader(objDb,GetRequest("CenterCd",""),GetRequest("YM",GetYM()),GetRequest("ptype","pTable"))%>
		</thead>
		<tbody>
			<%=MakeBody(objDb,GetRequest("CenterCd",""),GetRequest("JKubun",""),GetRequest("YM",GetYM()),GetRequest("ptype","pTable"))%>
		</tbody>
	</TABLE>
	</div>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = false;
		cpTblBtn.disabled = false;
		cpTblBtn.value = "���ʂ��R�s�[";
	//--></SCRIPT>
<%
		call closeDb(objDb)
	end if
%>
<%
	Call endHtml()
%>

<% sub	closeDb(objDb)
	objDb.Close
	set objDb = nothing
end sub %>

<% sub	endHtml() 	%>
	<!-- endHtml() start	-->
	</BODY>
	</HTML>
	<!-- endHtml() end		-->
<% end sub			%>

<%
'-------------------------------------------------------------
'Table Caption
'-------------------------------------------------------------
Function GetCaption(byval strYm,byval strCenterCD,byval strJKubun,byval strPType)
	dim	strCaption

	strCaption = ""
	select case strPType
	case "pList"
		strCaption = "�f�[�^�ꗗ"
	case "pTable"
		strCaption = GetPeriod(strYM) & "�� " & Right(strYM,2) & "���� �o�c�T����"
	case "pTableSyushi"
		strCaption = GetPeriod(strYM) & "�� " & Right(strYM,2) & "���� �o�c�T����+���x��"
	case "pTableJK"
		strCaption = GetPeriod(strYM) & "�� " & Right(strYM,2) & "���� ���Ƌ敪�T����"
	case "pTableJKKan"
		strCaption = GetPeriod(strYM) & "�� " & Right(strYM,2) & "���� ���Ƌ敪�T����(���Ԏ��ԃ`�F�b�N)"
	case "pTableJM"
		strCaption = GetPeriod(strYM) & "�� " & Right(strYM,2) & "���� ���Ƌ敪�ʓ����ڍ�"
	case "pTableJKYear"
		strCaption = "�N�Ԏ��Ƌ敪�ʊT����"
	case "pTableJKYearMonth","pTableJKYearMonth2"
		strCaption = "�N�Ԏ��Ƌ敪�ʊT����(����)"
	end select
	strCaption = strCaption & " " & strYm
	strCaption = strCaption & " " & GetCenterName(strCenterCD)
	strCaption = strCaption & " " & strJKubun
	dim	strSyushi
	strSyushi = GetRequest("SyushiCD","")
	if strSyushi <> "" then
		strCaption = strCaption & " " & strSyushi
	end if
	GetCaption = RTrim(strCaption)
End Function

'-------------------------------------------------------------
'�e�[�u�����e
'-------------------------------------------------------------
Function MakeBody(byVal objDb,byVal strCenterCD,byval strJKubun,byVal strYM,byval strTableType)
	dim	strHTML
	dim	objRs
	dim	strPersonName
	dim	strPersonCD
	dim	i
	dim	iKubun
	dim	aryWorkTM
	dim	strSql
	dim	errNumber
	dim	lngTotalTM
	dim	lngWorkTM

	strHTML = vbCrLf
	strHTML = strHTML & "<!-- MakeBody(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strTableType & ")-->" & vbCrLf

	'------------------------------------------------------------------------------
	'���R�[�h���e����HTML���쐬
	'------------------------------------------------------------------------------
	select case strTableType
	case "pList"
		strHTML = strHTML & GetTdList(objDb,strCenterCD,strJKubun,strYM,strTableType)
	case "pTable","pTableSyushi"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">����</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"����")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
'		strHTML = strHTML & "<TH rowspan=""4"">���㌴��</TH>"
		strHTML = strHTML & "<TH rowspan=""5"">����</TH>"
		strHTML = strHTML & "<TH>���ޔ�</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"���ޔ�")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>�H���d��</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�H���d��")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>���̑�</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"���̑��d��")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TH>���ڐl����</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"���ڐl����")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>���v</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�d��")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"����")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">���E���v</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�e���v")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"���E���v")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""5"">�Œ��</TH>"
		strHTML = strHTML & "<TH>�Ԑڐl����</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�Ԑڐl����")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>�ʏ�Ǘ���</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�ʏ�Ǘ���")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>���ʊǗ���</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"���ʊǗ���")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>�V�X�e����</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�V�X�e����")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>���v</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�o��")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�Œ��")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">�c�Ɨ��v</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"�c�Ɨ��v")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" rowspan=""2"">"
		if strCenterCD = "A" then
			strHTML = strHTML & "�Z���^�[��"
		else
			select case strTableType
			case "pTable"
				strHTML = strHTML & "���Ƌ敪��"
			case "pTableSyushi"
				strHTML = strHTML & "���x��"
			end select
		end if
		strHTML = strHTML & "</TH>"
		strHTML = strHTML & "<TH colspan=""2"">����</TH>"
		strHTML = strHTML & "<TH colspan=""2"">�v�捷</TH>"
		strHTML = strHTML & "<TH colspan=""2"">�݌v</TH>"
		strHTML = strHTML & "<TH colspan=""2"">�v�捷</TH>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "</TR>"

		if strCenterCD = "A" then
			' B:����PC/D:���ꕨ��/E:����PC/F:�ޗ�/G:���PC/H:�܈�PC/I:�L��
			dim	strC
			for each strC in Array("B","E","H","D","G","F","I")
				strHTML = strHTML & "<TR>"
				strHTML = strHTML & "<TD colspan=""2"">" & GetCenterName(strC) & "</TD>"
				strHTML = strHTML & GetTdValue(objDb,strC,strJKubun,strYM,"_")
				strHTML = strHTML & "</TR>"
			next
		else
			select case strTableType
			case "pTable"
				strSql = "select distinct JigyoKubunName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by JigyoKubunName"
				set objRs = objDb.Execute(strSql)
				do while objRs.Eof = False
					strHTML = strHTML & "<TR>"
					strHTML = strHTML & "<TD colspan=""2"">" & GetJKubunLink(strYm,strCenterCD,GetFields(objRs,"JigyoKubunName")) & "</TD>"
					strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"_" & GetFields(objRs,"JigyoKubunName"))
					strHTML = strHTML & "</TR>"
					objRs.MoveNext
				loop
					strHTML = strHTML & "<TR>"
					strHTML = strHTML & "<TD colspan=""2"">" & GetJKubunLink(strYm,strCenterCD,"_���̑�") & "</TD>"
					strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"__���̑�")
					strHTML = strHTML & "</TR>"
			case "pTableSyushi"
				strSql = "select distinct SyushiCD,SyushiName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by SyushiCD"
				set objRs = objDb.Execute(strSql)
				do while objRs.Eof = False
					strHTML = strHTML & vbCrLf & "<TR>"
					strHTML = strHTML & vbCrLf & "<TD colspan=""2"">" & GetSyushiLink(strYm,strCenterCD,GetFields(objRs,"SyushiCD"),GetFields(objRs,"SyushiName")) & "</TD>"
					strHTML = strHTML & vbCrLf & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"_Syushi_" & GetFields(objRs,"SyushiCD"))
					strHTML = strHTML & vbCrLf & "</TR>"
					objRs.MoveNext
				loop
			end select
		end if
	case "pTableJM"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" rowspan=""2"">���Ƌ敪��</TH>"
		strHTML = strHTML & "<TH colspan=""4"">����</TH>"
		strHTML = strHTML & "<TH colspan=""4"">�v��</TH>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "<TH>���ڐl����</TH>"
		strHTML = strHTML & "<TH>�Ԑڐl����</TH>"
		strHTML = strHTML & "<TH>����</TH>"
		strHTML = strHTML & "<TH>���v</TH>"
		strHTML = strHTML & "<TH>���ڐl����</TH>"
		strHTML = strHTML & "<TH>�Ԑڐl����</TH>"
		strHTML = strHTML & "</TR>"

		strSql = "select distinct JigyoKubunName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by JigyoKubunName"
		set objRs = objDb.Execute(strSql)
		do while objRs.Eof = False
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TD colspan=""1"">" & GetJKubunLink(strYm,strCenterCD,GetFields(objRs,"JigyoKubunName")) & "</TD>"
			strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"JM_" & GetFields(objRs,"JigyoKubunName"))
			strHTML = strHTML & "</TR>"
			objRs.MoveNext
		loop
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TD colspan=""1"">" & GetJKubunLink(strYm,strCenterCD,"_���̑�") & "</TD>"
			strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"JM_���̑�")
			strHTML = strHTML & "</TR>"
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3","pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"	' �N�Ԏ��Ƌ敪�ʊT����
		dim	strJK
		select case strTableType
		case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
			strJK = "CK"
		case "pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"	' �N�Ԏ��Ƌ敪�ʊT����
			strJK = "YK"
		end select
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "����")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""3"">����</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>���㌴��</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���㌴��")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>���ڐl����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���ڐl����")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�v</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "����")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" nowrap>���E���v</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���E���v")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""5"">�Œ��</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�Ԑڐl����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�Ԑڐl����")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�ʏ�Ǘ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�ʏ�Ǘ���")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>���ʊǗ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���ʊǗ���")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�V�X�e����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�V�X�e����")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�v</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�Œ��")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" nowrap>�c�Ɨ��v</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�c�Ɨ��v")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""6"">���ڍ�Ƃg</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�Ζ�����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�Ζ�����")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>��Ǝ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "��Ǝ���")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>���Ǝ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���Ǝ���")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�L������</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�L������")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>����H��</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "����H��")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�H��(�]�T����)</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�H��(�]�T����)")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""4"">�Ԑڍ�Ƃg</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�Ζ�����</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�Ζ�����(��)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>��Ǝ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "��Ǝ���(��)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>���Ǝ���</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���Ǝ���(��)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>�L������</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�L������(��)")
		strHTML = strHTML & "</TR>" & vbCrLf
		'-----------------------------------------------------------------------
		if strJK = "YK" then
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>�Ζ�����</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�Ζ�����(�v)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>��Ǝ���</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "��Ǝ���(�v)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>���Ǝ���</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "���Ǝ���(�v)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>�L������</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "�L������(�v)")
			strHTML = strHTML & "</TR>" & vbCrLf
		end if
	case "pTableJK","pTableJKKan"
		'-----------------------------------------------------------------------
		' pTableJK		���Ƌ敪�T����
		' pTableJKKan	���Ƌ敪�T����(�Ԑڂ���)
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">����</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK����")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH rowspan=""3"">����</TH>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">���㌴��</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���㌴��")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">���ڐl����</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���ڐl����")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">�v</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK����")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">���E���v</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���E���v")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH rowspan=""5"">�Œ��</TH>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">�Ԑڐl����</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�Ԑڐl����")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">�ʏ�Ǘ���</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�ʏ�Ǘ���")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">���ʊǗ���</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���ʊǗ���")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">�V�X�e����</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�V�X�e����")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">�v</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�Œ��")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">�c�Ɨ��v</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�c�Ɨ��v")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		for i = 1 to 18
			strHTML = strHTML & "<TD></TD>"
		next
		strHTML = strHTML & vbCrlf
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TD rowspan=""6"">���ڍ�Ƃg</TD>" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">�Ζ�����</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�Ζ�����")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">��Ǝ���</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK��Ǝ���")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">���Ǝ���</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���Ǝ���")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">�L������</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�L������")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightcyan"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">����H��</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK����H��")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightcyan"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">�H��(�]�T����)</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�H��(�]�T����)")
		strHTML = strHTML & "</TR>" & vbCrlf
		if strTableType = "pTableJKKan" then
			'-----------------------------------------------------------------------
			strHTML = strHTML & "<TR>" & vbCrlf
			strHTML = strHTML & "<TD rowspan=""4"">�Ԑڍ�Ƃg</TD>" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">�Ζ�����</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�Ζ�����(��)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">��Ǝ���</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK��Ǝ���(��)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">���Ǝ���</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK���Ǝ���(��)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">�L������</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK�L������(��)")
			strHTML = strHTML & "</TR>" & vbCrlf
			'--------------------------------------------
			'�e�펞�ԓ���
			'--------------------------------------------
			strHTML = strHTML & GetTdList(objDb,strCenterCD,strJKubun,strYM,"JK���ԓ���")
		end if
	end select
	set objRs = Nothing
	strHTML = strHTML & "<!-- MakeBody() End -->" & vbCrLf
	MakeBody = strHTML
End Function

Function GetJKubunLink(byval strYm,byval strCenterCD,byval strJKubun)
	dim	strLink
	strLink = "<A href=""" & Request.ServerVariables("URL")
	strLink = strLink & "?YM=" & strYm
	strLink = strLink & "&CenterCD=" & strCenterCD
	strLink = strLink & "&JKubun=" & Server.URLEncode(strJKubun)
	strLink = strLink & "&ptype=pTableJK"
	strLink = strLink & "&submit1=0"
	strLink = strLink & """>"
	strLink = strLink & strJKubun
	strLink = strLink & "</A>"
	GetJKubunLink = strLink
End Function

Function GetSyushiLink(byval strYm,byval strCenterCD,byval strSyushiCD,byval strSyushiName)
	dim	strLink
	strLink = "<A href=""" & Request.ServerVariables("URL")
	strLink = strLink & "?YM=" & strYm
	strLink = strLink & "&CenterCD=" & strCenterCD
	strLink = strLink & "&SyushiCD=" & Server.URLEncode(strSyushiCD)
	strLink = strLink & "&ptype=pTableJK"
	strLink = strLink & "&submit1=0"
	strLink = strLink & """>"
	strLink = strLink & strSyushiCD & " " & strSyushiName
	strLink = strLink & "</A>"
	GetSyushiLink = strLink
End Function
'------------------------------------------------------------------------------
'���Ƌ敪�T��(����)
'------------------------------------------------------------------------------
dim	objRsYK
Function GetTdValueYK(objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strYM2
	dim	strSql
	GetTdValueYK = ""
	dim	strHTML
	strHTML = ""
	select case strKubun
	case "YK����"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' �N�Ԏ��Ƌ敪�ʊT����
			strYM2 = clng(strYM) + 100
		end select
		strSql = YKSub(strCenterCD,strYM,strYM2,strJKubun,strKubun)
objDb.commandTimeout=600
		on error resume next
			set objRsYK = objDb.Execute(strSql)
			if Err.Number <> 0 then
				strHTML = MakeError(Err)
			end if
		on error goto 0
	case "YK���㌴��","YK���ڐl����","YK����","YK���E���v","YK�Ԑڐl����","YK�ʏ�Ǘ���","YK���ʊǗ���","YK�V�X�e����","YK�Œ��","YK�c�Ɨ��v"
	case "YK�Ζ�����"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' �N�Ԏ��Ƌ敪�ʊT����
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"1")
'		on error resume next
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
'			if Err.Number <> 0 then
'				strHTML = MakeError(Err)
'			end if
'		on error goto 0
	case "YK�Ζ�����(��)"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' �N�Ԏ��Ƌ敪�ʊT����
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"2")
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
	case "YK�Ζ�����(�v)"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' �N�Ԏ��Ƌ敪�ʊT����
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"3")
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
	case "YK�Ζ�����(��)","YK�Ζ�����(�v)","YK��Ǝ���","YK��Ǝ���(��)","YK��Ǝ���(�v)","YK���Ǝ���","YK���Ǝ���(��)","YK���Ǝ���(�v)","YK�L������","YK�L������(��)","YK�L������(�v)","YK����H��","YK�H��(�]�T����)"
	case else
		exit function
	end select
	if strHTML <> "" then
		GetTdValueYK = strHTML
		exit function
	end if
	strHTML = vbCrLf & "<!--" & strCenterCD & "," & strJKubun & "," & strYM & "," & strYM2 & "," & strKubun & "-->" & vbCrLf
	strHTML = strHTML & "<!--" & strSql & "-->" & vbCrLf
	dim	lngRst
	dim	lngPln
	dim	lngRstSum
	dim	lngPlnSum
	dim	lngRstSumA
	dim	lngPlnSumA
	dim	strRst
	dim	strPln
	dim	i
	lngRstSum = 0
	lngPlnSum = 0
	lngRstSumA = 0
	lngPlnSumA = 0
	strRst = ""
	strPln = ""
	for i = 1 to 12
		dim	m
		m = i + 3
		if m > 12 then
			m = m - 12
		end if
		m = right("0" & m,2)
		select case strKubun
		case "YK����"
			lngRst = CDbl(objRsYK.Fields("ARst" & m))
			lngPln = CDbl(objRsYK.Fields("APln" & m))
		case "YK���㌴��"
			lngRst = CDbl(objRsYK.Fields("BRst" & m))
			lngPln = CDbl(objRsYK.Fields("BPln" & m))
		case "YK���ڐl����"
			lngRst = CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK����"
			lngRst = CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK���E���v"
			lngRst = CDbl(objRsYK.Fields("ARst" & m)) - (CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m)))
			lngPln = CDbl(objRsYK.Fields("APln" & m)) - (CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m)))
		case "YK�Ԑڐl����"
			lngRst = CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK�ʏ�Ǘ���"
			lngRst = CDbl(objRsYK.Fields("C2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C2Pln" & m))
		case "YK���ʊǗ���"
			lngRst = CDbl(objRsYK.Fields("C9Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C9Pln" & m))
		case "YK�V�X�e����"
			lngRst = CDbl(objRsYK.Fields("DRst" & m))
			lngPln = CDbl(objRsYK.Fields("DPln" & m))
		case "YK�Œ��"
			lngRst = CDbl(objRsYK.Fields("X2Rst" & m)) + CDbl(objRsYK.Fields("C2Rst" & m)) + CDbl(objRsYK.Fields("C9Rst" & m)) + CDbl(objRsYK.Fields("DRst" & m))
			lngPln = CDbl(objRsYK.Fields("X2Pln" & m)) + CDbl(objRsYK.Fields("C2Pln" & m)) + CDbl(objRsYK.Fields("C9Pln" & m)) + CDbl(objRsYK.Fields("DPln" & m))
		case "YK�c�Ɨ��v"
			lngRst = CDbl(objRsYK.Fields("ARst" & m)) - (CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m)))
			lngPln = CDbl(objRsYK.Fields("APln" & m)) - (CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m)))
			lngRst = lngRst - (CDbl(objRsYK.Fields("X2Rst" & m)) + CDbl(objRsYK.Fields("C2Rst" & m)) + CDbl(objRsYK.Fields("C9Rst" & m)) + CDbl(objRsYK.Fields("DRst" & m)))
			lngPln = lngPln - (CDbl(objRsYK.Fields("X2Pln" & m)) + CDbl(objRsYK.Fields("C2Pln" & m)) + CDbl(objRsYK.Fields("C9Pln" & m)) + CDbl(objRsYK.Fields("DPln" & m)))
		case "YK�Ζ�����"
			lngRst = GetFieldVal(objRsYK, "TM101Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM101Pln" & m)
		case "YK�Ζ�����(��)"
			lngRst = GetFieldVal(objRsYK, "TM102Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM102Pln" & m)
		case "YK�Ζ�����(�v)"
			lngRst = GetFieldVal(objRsYK, "ARst" & m)
			lngPln = GetFieldVal(objRsYK, "APln" & m)
		case "YK��Ǝ���"
			lngRst = GetFieldVal(objRsYK, "TM201Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM201Pln" & m)
		case "YK��Ǝ���(��)"
			lngRst = GetFieldVal(objRsYK, "TM202Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM202Pln" & m)
		case "YK��Ǝ���(�v)"
			lngRst = GetFieldVal(objRsYK, "TM201Rst" & m) + GetFieldVal(objRsYK, "TM202Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM201Pln" & m) + GetFieldVal(objRsYK, "TM202Pln" & m)
		case "YK���Ǝ���"
			lngRst = GetFieldVal(objRsYK, "TM301Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM301Pln" & m)
		case "YK���Ǝ���(��)"
			lngRst = GetFieldVal(objRsYK, "TM302Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM302Pln" & m)
		case "YK���Ǝ���(�v)"
			lngRst = GetFieldVal(objRsYK, "TM301Rst" & m) + GetFieldVal(objRsYK, "TM302Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM301Pln" & m) + GetFieldVal(objRsYK, "TM302Pln" & m)
		case "YK�L������"
			lngRst = GetFieldVal(objRsYK, "TM401Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM401Pln" & m)
		case "YK�L������(��)"
			lngRst = GetFieldVal(objRsYK, "TM402Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM402Pln" & m)
		case "YK�L������(�v)"
			lngRst = GetFieldVal(objRsYK, "TM401Rst" & m) + GetFieldVal(objRsYK, "TM402Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM401Pln" & m) + GetFieldVal(objRsYK, "TM402Pln" & m)
		case "YK����H��"
			lngRst = 0
			lngPln = 0
		case else
			lngRst = 0
			lngPln = 0
		end select
		lngRstSum = lngRstSum + lngRst
		lngPlnSum = lngPlnSum + lngPln
		lngRstSumA = lngRstSumA + GetFieldVal(objRsYK, "ARst" & m)
		lngPlnSumA = lngPlnSumA + GetFieldVal(objRsYK, "APln" & m)
		strRst = strRst & "<td class=""Number"">" & GetNumber(lngRst,"") & "</td>"
		strPln = strPln & "<td class=""Number"">" & GetNumber(lngPln,"") & "</td>"
	next
	dim	strRt
	strRt = ""
	if lngRstSumA <> 0 then
		strRt = GetNumber(lngRstSum/lngRstSumA*100,"")
	end if
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngRstSum,"") & "</td><td class=""Number"">" & strRt & "</td>" & strRst
	strRt = ""
	if lngPlnSumA <> 0 then
		strRt = GetNumber(lngPlnSum/lngPlnSumA*100,"")
	end if
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngPlnSum,"") & "</td><td class=""Number"">" & strRt & "</td>" & strPln
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngRstSum - lngPlnSum,"") & "</td>" & vbCrLf
	GetTdValueYK = strHTML
End Function

Function GetTdValueJK(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	GetTdValueJK = GetTdValueYK(objDb,strCenterCD,strJKubun,strYM,strKubun)
	if GetTdValueJK <> "" then
		exit function
	end if
	dim	strSql
	dim	objRs
	dim	strHTML
	'------------------------------------------------------------------------------
	'���Ƌ敪�T��(����)
	'------------------------------------------------------------------------------
	strHTML = ""
	strHTML = strHTML & "<!-- GetTdValueJK(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strKubun & ")-->" & vbCrLf
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	strHTML = strHTML & "<!-- " & vbCrLf & strSql & vbCrLf & ")-->" & vbCrLf
	if strSql <> "" then
		on error resume next
			set objRs = objDb.Execute(strSql)
			if Err.Number <> 0 then
				strHTML = MakeError(objErr)
			end if
			if objRs is Nothing then
				strHTML = strHTML & "<tr><td>" & objDb.Errors.Count & "</td>" & vbCrlf
		        Dim errX
		        For Each errX In objDb.Errors
					strHTML = strHTML & "<td>" & errX.Description & "</td>" & vbCrlf
		        Next
				strHTML = strHTML & "<td>" & strSql & "</td></tr>" & vbCrlf
'			End if
'			if not objRs is Nothing then
			else
				if objRs.Eof = False then
					dim	f
					for each f in objRs.Fields
						strHTML = strHTML & vbTab & "<td class=""Number"">" & GetNumber(GetFields(objRs,f.Name),"") & "</td>" & vbCrlf
					next
				end if
				set objRs = nothing
			end if
		on error goto 0
	end if
	strHTML = strHTML & "<!-- GetTdValueJK() End -->" & vbCrLf
	GetTdValueJK = strHTML
End Function


Function GetTdValue(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strSql
	dim	objRs
	dim	strHTML
	strHTML = ""
	'------------------------------------------------------------------------------
	'SQL���s
	'------------------------------------------------------------------------------
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			strHTML = MakeError(objErr)
		end if
		if objRs is Nothing then
			strHTML = strHTML & vbCrLf & "<tr><td>" & objDb.Errors.Count & "</td>"
	        Dim errX
	        For Each errX In objDb.Errors
				strHTML = strHTML & vbCrLf & "<td>" & errX.Description & "</td>"
	        Next
			strHTML = strHTML & vbCrLf & "<td>" & strSql & "</td></tr>"
		End if
	on error goto 0
	if strHTML = "" then
		strHTML = strHTML & vbCrLf & "<!-- "
		strHTML = strHTML & vbCrLf & strSql
		strHTML = strHTML & vbCrLf & " -->"
		if objRs.Eof = False then
			dim	f
			for each f in objRs.Fields
				strHTML = strHTML & vbCrLf & "<td class=""Number"">" & GetNumber(GetFields(objRs,f.Name),"") & "</td>"
			next
		end if
	end if
	GetTdValue = strHTML
End Function

Function GetTdList(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strSql
	dim	objRs
	dim	strHTML
	strHTML = ""
	'------------------------------------------------------------------------------
	'�f�[�^�ꗗ
	'------------------------------------------------------------------------------
	strHTML = ""
	strHTML = strHTML & "<!-- GetTdList(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strKubun & ")-->" & vbCrLf
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	strHTML = strHTML & "<!-- " & vbCrLf & strSql & vbCrLf & ")-->" & vbCrLf
	'------------------------------------------------------------------------------
	'SQL���s
	'------------------------------------------------------------------------------
	on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			strHTML = MakeError(objErr)
		end if
		if objRs is Nothing then
			strHTML = strHTML & "<tr><td>" & objDb.Errors.Count & "</td>"
	        Dim errX
	        For Each errX In objDb.Errors
				strHTML = strHTML & "<td>" & errX.Description & "</td>"
	        Next
			strHTML = strHTML & "<td>" & strSql & "</td></tr>"
		End if
	on error goto 0
	if not objRs is Nothing then
		dim	f
		if strJKubun = "Header" then
				strHTML = strHTML & "<TR>" & vbCrlf
				for each f in objRs.Fields
					strHTML = strHTML & "<th>" & f.Name & "</th>" & vbCrlf
				next
				strHTML = strHTML & "</TR>" & vbCrlf
		else
			do while objRs.Eof = False
				for each f in objRs.Fields
					strHTML = strHTML & "<!--bgcolor " & f & "-->" & vbCrlf
					select case RTrim(f)
					case "��Ǝ���(����)","��Ǝ���(�Ԑ�)"
						strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
					case "���Ǝ���(����)","���Ǝ���(�Ԑ�)"
						strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
					case "�L������(����)","�L������(�Ԑ�)"
						strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
					case else
						strHTML = strHTML & "<TR>" & vbCrlf
					end select
					exit for
				next
				for each f in objRs.Fields
'					strHTML = strHTML & "<td>" & GetFields(objRs,f.Name) & "</td>" & vbCrlf
					strHTML = strHTML & GetFieldTd(objRs,f) & vbCrlf
				next
				strHTML = strHTML & "</TR>" & vbCrlf
				objRs.MoveNext
			Loop
		end if
	end if
	strHTML = strHTML & "<!-- GetTdList() End -->" & vbCrLf
	GetTdList = strHTML
End Function
'-------------------------------------------------------------
'<td>��Ԃ�
'-------------------------------------------------------------
Function GetFieldTd(objRs,f)
	dim	strTd
	dim	strValue
	strValue = GetFields(objRs,f.Name)
	strTd = "<td"
	strTd = strTd & " title=""" & f.name & " " & f.type & """"
	select case f.type
	Case 2 , 3 , 5 , 6 ,131	' ���l(Integer)
		strTd = strTd & " class=""Number"""
		strValue = GetNumber(strValue,"")
	case else
		strTd = strTd & " class=""Character"""
	end select
'	strTd = strTd & " nowrap"
	strTd = strTd & ">"
	strTd = strTd & strValue
	strTd = strTd & "</td>"
	GetFieldTd = strTd
End Function
'-------------------------------------------------------------
'�z��N���A
'-------------------------------------------------------------
Function ClearArray(aryV())
	dim	i
	for i = lbound(aryV) to ubound(aryV)
		aryV(i) = 0
	next
	ClearArray = i
End Function

'-------------------------------------------------------------
'�z��̉��Z
'-------------------------------------------------------------
Function AddArray(aryV,objRs)
	dim	i
	dim	aryFld
	aryFld = Array("A_Prev","A_Plan","A_Result","A_Margin","T_Prev","T_Plan","T_Result","T_Margin")
	for i = lbound(aryFld) to ubound(aryFld)
		aryV(i) = aryV(i) + CLng(GetFields(objRs,aryFld(i)))
	next
	AddArray = i
End Function

'-------------------------------------------------------------
'�z���TD�v�f��Ԃ�
'-------------------------------------------------------------
Function GetTdArray(aryV())
	dim	i
	dim	strTd
	strTd = ""
	for i = lbound(aryV) to ubound(aryV)
		strTd = strTd & "<TD>" & GetNumber(aryV(i),"") & "/TD" & vbcrlf
	next
	GetTdArray = strTd
End Function


'-------------------------------------------------------------
'Field�̒l��Ԃ�
'-------------------------------------------------------------
Function GetFields(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	if strFieldName = "" then
		v = 0
	else
		select case strFieldName
		case "Y_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"Y_Prev"),GetFieldValue(objRs,"Y_Prev_Hi"))
		case "Y_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"Y_Plan"),GetFieldValue(objRs,"Y_Plan_Hi"))
'		case "Y_Margin"
'			v = CLng(GetFields(objRs,"Y_Plan")) - CLng(GetFields(objRs,"Y_Prev"))
		case "A_Result_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Result"),GetFieldValue(objRs,"A_Result_Hi"))
		case "T_Result_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Result"),GetFieldValue(objRs,"T_Result_Hi"))
		case "A_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Plan"),GetFieldValue(objRs,"A_Plan_Hi"))
		case "T_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Plan"),GetFieldValue(objRs,"T_Plan_Hi"))
		case "A_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Prev"),GetFieldValue(objRs,"A_Prev_Hi"))
		case "T_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Prev"),GetFieldValue(objRs,"T_Prev_Hi"))
		case "A_Prev_Margin"
			v = CLng(GetFields(objRs,"A_Result")) - CLng(GetFields(objRs,"A_Prev"))
		case "T_Prev_Margin"
			v = CLng(GetFields(objRs,"T_Result")) - CLng(GetFields(objRs,"T_Prev"))
		case "A_Margin"
			v = CLng(GetFields(objRs,"A_Result")) - CLng(GetFields(objRs,"A_Plan"))
		case "T_Margin"
			v = CLng(GetFields(objRs,"T_Result")) - CLng(GetFields(objRs,"T_Plan"))
		case "A_UriSa"
			v = CLng(GetFieldValue(objRs,"A_Uri")) - CLng(GetFieldValue(objRs,"A_UriSa"))
		case "A_RiekiSa"
			v = CLng(GetFieldValue(objRs,"A_Rieki")) - CLng(GetFieldValue(objRs,"A_RiekiSa"))
		case "T_UriSa"
			v = CLng(GetFieldValue(objRs,"T_Uri")) - CLng(GetFieldValue(objRs,"T_UriSa"))
		case "T_RiekiSa"
			v = CLng(GetFieldValue(objRs,"T_Rieki")) - CLng(GetFieldValue(objRs,"T_RiekiSa"))
		case else
			v = RTrim(objRs.Fields(strFieldName))
			if isnull(v) then
				v = 0
			end if
		end select
	end if
	GetFields = v
End Function

Function GetPercent(byval v1,byval v2)
	dim	v
	v = 0
	if clng(v2) <> 0 then
		v = clng(v1) * 100 / clng(v2)
		v = Round(v,0)
	end if
	GetPercent = v
End Function

Function GetFieldVal(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	v = GetFieldValue(objRs, strFieldName)
	GetFieldVal = cdbl(v)
End Function

Function GetFieldValue(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	if strFieldName = "" then
		v = 0
	else
		v = objRs.Fields(strFieldName)
		if isnull(v) then
			v = 0
		end if
	end if
	GetFieldValue = v
End Function
'-------------------------------------------------------------
'������Ԃ�
'-------------------------------------------------------------
Function GetNumber(byVal v,byVal strFormat)
	dim	strNumber

	strNumber = ""
	if isnull(v) = False then
		select case strFormat
		case "%"
			if CLng(v) <> 0 then
				strNumber = formatnumber(v,0,,,-1) & "%"
			end if
		case else
			if CLng(v) <> 0 then
				strNumber = formatnumber(v,0,,,-1)
			end if
		end select
	end if
	GetNumber = strNumber
End Function

Function GetFromA(byVal strYM,byVal strCenterCD)
	GetFromA = " "
	GetFromA = GetFromA & vbcrlf & "from (select"
	GetFromA = GetFromA & vbcrlf & " sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Prev,0)) Prev_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Prev,0)) Prev_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Prev,0)) Prev_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Prev,0)) Prev_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Prev,0)) Prev_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Prev,0)) Prev_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Prev,0)) Prev_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Prev,0)) Prev_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Plan,0)) Plan_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Plan,0)) Plan_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Plan,0)) Plan_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Plan,0)) Plan_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Plan,0)) Plan_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Plan,0)) Plan_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Plan,0)) Plan_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Plan,0)) Plan_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Result,0)) Result_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Result,0)) Result_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Result,0)) Result_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Result,0)) Result_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Result,0)) Result_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Result,0)) Result_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Result,0)) Result_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Result,0)) Result_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Prev,0)) tPrev_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Prev,0)) tPrev_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Prev,0)) tPrev_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Prev,0)) tPrev_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Prev,0)) tPrev_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Prev,0)) tPrev_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Prev,0)) tPrev_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Prev,0)) tPrev_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Plan,0)) tPlan_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Plan,0)) tPlan_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Plan,0)) tPlan_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Plan,0)) tPlan_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Plan,0)) tPlan_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Plan,0)) tPlan_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Plan,0)) tPlan_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Plan,0)) tPlan_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Result,0)) tResult_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Result,0)) tResult_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Result,0)) tResult_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Result,0)) tResult_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Result,0)) tResult_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Result,0)) tResult_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Result,0)) tResult_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Result,0)) tResult_X2"

	GetFromA = GetFromA & vbcrlf & " from IrData"
	GetFromA = GetFromA & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
	GetFromA = GetFromA & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
	GetFromA = GetFromA & vbcrlf & ") a"
End Function

'-------------------------------------------------------------
'SQL where ����
'-------------------------------------------------------------
Function SqlWhere(byVal strAnd, byVal strField, byVal strValue)
	SqlWhere = ""
	if strValue = "" then
		exit function
	end if
	dim	strCmp
	strCmp = "="
	if left(strValue,1) = "-" then
		strCmp = "<>"
		strValue = right(strValue,len(strValue)-1)
	end if
	if instr(strValue,"%") > 0 then
		if strCmp = "=" then
			strCmp = "like"
		else
			strCmp = "not like"
		end if
	elseif instr(strValue,",") > 0 then
		if strCmp = "=" then
			strCmp = "in"
		else
			strCmp = "not in"
		end if
		strValue = "('" & replace(strValue, ",", "','") & "')"
	end if
	SqlWhere = vbCrLf & strAnd & " " & strField & " " & strCmp & " " & strValue
End Function
'-------------------------------------------------------------
'����SQL
'-------------------------------------------------------------
Function MakeSql(byVal strCenterCD,byval strJKubun,byVal strYM,byval strKubun)
	dim	strSql
	dim	i
	dim	strSqlAdd

	strSql		= ""
	strSqlAdd	= ""

	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if

	strKubun = rtrim(strKubun)
	if strKubun = "pList" then
		strSql = "select"
'YM 	CenterCD 	SyushiCD 	KamokuCD 	Plan 	Result 	Prev
		strSql = strSql & vbcrlf & " i.CenterCD ""�Z���^�["""
		strSql = strSql & vbcrlf & ",i.YM ""�N��"""
'		strSql = strSql & vbcrlf & ",i.KamokuCD ""�Ȗ�"""
		strSql = strSql & vbcrlf & ",i.KamokuCD + RTrim(' ' + ifnull(k.KamokuName,'')) ""�Ȗ�"""
		strSql = strSql & vbcrlf & ",ifnull(j.JigyoKubunName,'') ""���Ƌ敪"""
		strSql = strSql & vbcrlf & ",i.SyushiCD ""���x"""
		strSql = strSql & vbcrlf & ",i.Result ""����"""
		strSql = strSql & vbcrlf & ",i.Plan ""�v��"""
		strSql = strSql & vbcrlf & ",i.Prev ""�O�N"""
		strSql = strSql & vbcrlf & " from IrData i"
		strSql = strSql & vbcrlf & " left outer join Kamoku k on (i.KamokuCD = k.KamokuCD)"
		strSql = strSql & vbcrlf & " left outer join JigyoKubun j on (i.CenterCD = j.CenterCD and i.SyushiCD = j.SyushiCD)"
		strSql = strSql & vbcrlf & " WHERE i.YM = '" & strYM &"'"
		strSql = strSql & vbcrlf & "   AND i.CenterCD = '" & strCenterCD &"'"
		strSql = strSql & vbcrlf & "   AND (i.Result <> 0 or i.Plan <> 0 or i.Prev <> 0)"
		strSql = strSql & SqlWhere("and", "i.SyushiCD", strSyushiCD)
'		if strSyushiCD <> "" then
'			strSql = strSql & vbcrlf & "   AND i.SyushiCD = '" & strSyushiCD & "'"
'		end if
		if left(strKubun,8) <> "_Syushi_" then
			if strJKubun = "_���̑�" then
				strSql = strSql & vbcrlf & "   AND i.SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
			elseif strJKubun <> "" then
				strSql = strSql & vbcrlf & "   AND i.SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
			end if
		end if
'		if strSyushiCD <> "" then
'			strSql = strSql & vbcrlf & "   AND i.SyushiCd = '" & strSyushiCD & "'"
'		end if
		strSql = strSql & vbcrlf & " order by i.YM,i.CenterCD,i.KamokuCD,""���Ƌ敪"",i.SyushiCD"
	else
		if left(strKubun,3) = "JM_" then
			strKubun = right(strKubun,len(strKubun) - 3)
			strSql = "select"
			strSql = strSql & vbcrlf & " sum((Result * if(KamokuCD like 'A%',1,0))) R_Uri"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1,if(KamokuCD like 'B%',-1,0)))) R_Rieki"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD in ('C0100','C0200','C0300','C0400'),1,if(KamokuCD in ('X0200'),-1,0)))) R_Choku"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD in ('X0200'),1,0))) R_Kan"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD like 'A%',1,0))) P_Uri"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD like 'A%',1,if(KamokuCD like 'B%',-1,0)))) P_Rieki"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD in ('C0100','C0200','C0300','C0400'),1,if(KamokuCD in ('X0200'),-1,0)))) P_Choku"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD in ('X0200'),1,0))) P_Kan"
			strSql = strSql & vbcrlf & " from IrData"
			strSql = strSql & vbcrlf & " WHERE YM = '" & strYM &"'"
			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
			strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'X9999'"
			if left(strKubun,8) <> "_Syushi_" then
				if strKubun = "_���̑�" then
					strSql = strSql & vbcrlf & "   AND SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
				else
					strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strKubun & "')"
				end if
			end if
			strSql = strSql & SqlWhere("and", "SyushiCd", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		elseif left(strKubun,1) = "_" then
			strKubun = right(strKubun,len(strKubun) - 1)
			strSql = "select"
			strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1, 0),0)) A_Uri"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1,-1),0)) A_Rieki"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1, 0),0)) A_UriSa"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1,-1),0)) A_RiekiSa"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1, 0))) T_Uri"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1,-1))) T_Rieki"
			strSql = strSql & vbcrlf & ",sum((Plan * if(KamokuCD like 'A%',1, 0))) T_UriSa"
			strSql = strSql & vbcrlf & ",sum((Plan * if(KamokuCD like 'A%',1,-1))) T_RiekiSa"
			strSql = strSql & vbcrlf & " from IrData"
			strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
			strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			if left(strKubun,7) <> "Syushi_" then
				if strKubun = "_���̑�" then
					strSql = strSql & vbcrlf & "   AND SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
				elseif strKubun <> "" then
					strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strKubun & "')"
				end if
			end if
			strSql = strSql & SqlWhere("and", "SyushiCd", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		elseif left(strKubun,2) = "YK" then
			dim	strYM2
			strYM2 = strYM
			select case GetRequest("ptype","pTable")	'strTableType
			case "pTableJKYearMonth"	' �N�Ԏ��Ƌ敪�ʊT����
				strYM2 = clng(strYM) + 100
			end select
			select case strKubun
			case "YK����"
				strSql = "select"
				strSql = strSql & vbcrlf & " (ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",ARst04"
				strSql = strSql & vbcrlf & ",ARst05"
				strSql = strSql & vbcrlf & ",ARst06"
				strSql = strSql & vbcrlf & ",ARst07"
				strSql = strSql & vbcrlf & ",ARst08"
				strSql = strSql & vbcrlf & ",ARst09"
				strSql = strSql & vbcrlf & ",ARst10"
				strSql = strSql & vbcrlf & ",ARst11"
				strSql = strSql & vbcrlf & ",ARst12"
				strSql = strSql & vbcrlf & ",ARst01"
				strSql = strSql & vbcrlf & ",ARst02"
				strSql = strSql & vbcrlf & ",ARst03"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",APln04"
				strSql = strSql & vbcrlf & ",APln05"
				strSql = strSql & vbcrlf & ",APln06"
				strSql = strSql & vbcrlf & ",APln07"
				strSql = strSql & vbcrlf & ",APln08"
				strSql = strSql & vbcrlf & ",APln09"
				strSql = strSql & vbcrlf & ",APln10"
				strSql = strSql & vbcrlf & ",APln11"
				strSql = strSql & vbcrlf & ",APln12"
				strSql = strSql & vbcrlf & ",APln01"
				strSql = strSql & vbcrlf & ",APln02"
				strSql = strSql & vbcrlf & ",APln03"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK���㌴��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",BRst04"
				strSql = strSql & vbcrlf & ",BRst05"
				strSql = strSql & vbcrlf & ",BRst06"
				strSql = strSql & vbcrlf & ",BRst07"
				strSql = strSql & vbcrlf & ",BRst08"
				strSql = strSql & vbcrlf & ",BRst09"
				strSql = strSql & vbcrlf & ",BRst10"
				strSql = strSql & vbcrlf & ",BRst11"
				strSql = strSql & vbcrlf & ",BRst12"
				strSql = strSql & vbcrlf & ",BRst01"
				strSql = strSql & vbcrlf & ",BRst02"
				strSql = strSql & vbcrlf & ",BRst03"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",BPln04"
				strSql = strSql & vbcrlf & ",BPln05"
				strSql = strSql & vbcrlf & ",BPln06"
				strSql = strSql & vbcrlf & ",BPln07"
				strSql = strSql & vbcrlf & ",BPln08"
				strSql = strSql & vbcrlf & ",BPln09"
				strSql = strSql & vbcrlf & ",BPln10"
				strSql = strSql & vbcrlf & ",BPln11"
				strSql = strSql & vbcrlf & ",BPln12"
				strSql = strSql & vbcrlf & ",BPln01"
				strSql = strSql & vbcrlf & ",BPln02"
				strSql = strSql & vbcrlf & ",BPln03"
				strSql = strSql & vbcrlf & ",(BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"

			case "YK���ڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C1Rst04-X2Rst04"
				strSql = strSql & vbcrlf & ",C1Rst05-X2Rst05"
				strSql = strSql & vbcrlf & ",C1Rst06-X2Rst06"
				strSql = strSql & vbcrlf & ",C1Rst07-X2Rst07"
				strSql = strSql & vbcrlf & ",C1Rst08-X2Rst08"
				strSql = strSql & vbcrlf & ",C1Rst09-X2Rst09"
				strSql = strSql & vbcrlf & ",C1Rst10-X2Rst10"
				strSql = strSql & vbcrlf & ",C1Rst11-X2Rst11"
				strSql = strSql & vbcrlf & ",C1Rst12-X2Rst12"
				strSql = strSql & vbcrlf & ",C1Rst01-X2Rst01"
				strSql = strSql & vbcrlf & ",C1Rst02-X2Rst02"
				strSql = strSql & vbcrlf & ",C1Rst03-X2Rst03"
				strSql = strSql & vbcrlf & ",(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C1Pln04-X2Pln04"
				strSql = strSql & vbcrlf & ",C1Pln05-X2Pln05"
				strSql = strSql & vbcrlf & ",C1Pln06-X2Pln06"
				strSql = strSql & vbcrlf & ",C1Pln07-X2Pln07"
				strSql = strSql & vbcrlf & ",C1Pln08-X2Pln08"
				strSql = strSql & vbcrlf & ",C1Pln09-X2Pln09"
				strSql = strSql & vbcrlf & ",C1Pln10-X2Pln10"
				strSql = strSql & vbcrlf & ",C1Pln11-X2Pln11"
				strSql = strSql & vbcrlf & ",C1Pln12-X2Pln12"
				strSql = strSql & vbcrlf & ",C1Pln01-X2Pln01"
				strSql = strSql & vbcrlf & ",C1Pln02-X2Pln02"
				strSql = strSql & vbcrlf & ",C1Pln03-X2Pln03"
				strSql = strSql & vbcrlf & ",((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "+(C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",BRst04+C1Rst04-X2Rst04"
				strSql = strSql & vbcrlf & ",BRst05+C1Rst05-X2Rst05"
				strSql = strSql & vbcrlf & ",BRst06+C1Rst06-X2Rst06"
				strSql = strSql & vbcrlf & ",BRst07+C1Rst07-X2Rst07"
				strSql = strSql & vbcrlf & ",BRst08+C1Rst08-X2Rst08"
				strSql = strSql & vbcrlf & ",BRst09+C1Rst09-X2Rst09"
				strSql = strSql & vbcrlf & ",BRst10+C1Rst10-X2Rst10"
				strSql = strSql & vbcrlf & ",BRst11+C1Rst11-X2Rst11"
				strSql = strSql & vbcrlf & ",BRst12+C1Rst12-X2Rst12"
				strSql = strSql & vbcrlf & ",BRst01+C1Rst01-X2Rst01"
				strSql = strSql & vbcrlf & ",BRst02+C1Rst02-X2Rst02"
				strSql = strSql & vbcrlf & ",BRst03+C1Rst03-X2Rst03"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03)"
				strSql = strSql & vbcrlf & "+(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",BPln04+C1Pln04-X2Pln04"
				strSql = strSql & vbcrlf & ",BPln05+C1Pln05-X2Pln05"
				strSql = strSql & vbcrlf & ",BPln06+C1Pln06-X2Pln06"
				strSql = strSql & vbcrlf & ",BPln07+C1Pln07-X2Pln07"
				strSql = strSql & vbcrlf & ",BPln08+C1Pln08-X2Pln08"
				strSql = strSql & vbcrlf & ",BPln09+C1Pln09-X2Pln09"
				strSql = strSql & vbcrlf & ",BPln10+C1Pln10-X2Pln10"
				strSql = strSql & vbcrlf & ",BPln11+C1Pln11-X2Pln11"
				strSql = strSql & vbcrlf & ",BPln12+C1Pln12-X2Pln12"
				strSql = strSql & vbcrlf & ",BPln01+C1Pln01-X2Pln01"
				strSql = strSql & vbcrlf & ",BPln02+C1Pln02-X2Pln02"
				strSql = strSql & vbcrlf & ",BPln03+C1Pln03-X2Pln03"
				strSql = strSql & vbcrlf & ",((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03))"
				strSql = strSql & vbcrlf & "+((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK���E���v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum((if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "+(C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",ARst04-(BRst04+C1Rst04-X2Rst04)"
				strSql = strSql & vbcrlf & ",ARst05-(BRst05+C1Rst05-X2Rst05)"
				strSql = strSql & vbcrlf & ",ARst06-(BRst06+C1Rst06-X2Rst06)"
				strSql = strSql & vbcrlf & ",ARst07-(BRst07+C1Rst07-X2Rst07)"
				strSql = strSql & vbcrlf & ",ARst08-(BRst08+C1Rst08-X2Rst08)"
				strSql = strSql & vbcrlf & ",ARst09-(BRst09+C1Rst09-X2Rst09)"
				strSql = strSql & vbcrlf & ",ARst10-(BRst10+C1Rst10-X2Rst10)"
				strSql = strSql & vbcrlf & ",ARst11-(BRst11+C1Rst11-X2Rst11)"
				strSql = strSql & vbcrlf & ",ARst12-(BRst12+C1Rst12-X2Rst12)"
				strSql = strSql & vbcrlf & ",ARst01-(BRst01+C1Rst01-X2Rst01)"
				strSql = strSql & vbcrlf & ",ARst02-(BRst02+C1Rst02-X2Rst02)"
				strSql = strSql & vbcrlf & ",ARst03-(BRst03+C1Rst03-X2Rst03)"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03)"
				strSql = strSql & vbcrlf & "-((BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03)"
				strSql = strSql & vbcrlf & "+(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",APln04-(BPln04+C1Pln04-X2Pln04)"
				strSql = strSql & vbcrlf & ",APln05-(BPln05+C1Pln05-X2Pln05)"
				strSql = strSql & vbcrlf & ",APln06-(BPln06+C1Pln06-X2Pln06)"
				strSql = strSql & vbcrlf & ",APln07-(BPln07+C1Pln07-X2Pln07)"
				strSql = strSql & vbcrlf & ",APln08-(BPln08+C1Pln08-X2Pln08)"
				strSql = strSql & vbcrlf & ",APln09-(BPln09+C1Pln09-X2Pln09)"
				strSql = strSql & vbcrlf & ",APln10-(BPln10+C1Pln10-X2Pln10)"
				strSql = strSql & vbcrlf & ",APln11-(BPln11+C1Pln11-X2Pln11)"
				strSql = strSql & vbcrlf & ",APln12-(BPln12+C1Pln12-X2Pln12)"
				strSql = strSql & vbcrlf & ",APln01-(BPln01+C1Pln01-X2Pln01)"
				strSql = strSql & vbcrlf & ",APln02-(BPln02+C1Pln02-X2Pln02)"
				strSql = strSql & vbcrlf & ",APln03-(BPln03+C1Pln03-X2Pln03)"
				strSql = strSql & vbcrlf & ",((ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03))"
				strSql = strSql & vbcrlf & "-(((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03))"
				strSql = strSql & vbcrlf & "+((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03))) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK�Ԑڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",X2Rst04"
				strSql = strSql & vbcrlf & ",X2Rst05"
				strSql = strSql & vbcrlf & ",X2Rst06"
				strSql = strSql & vbcrlf & ",X2Rst07"
				strSql = strSql & vbcrlf & ",X2Rst08"
				strSql = strSql & vbcrlf & ",X2Rst09"
				strSql = strSql & vbcrlf & ",X2Rst10"
				strSql = strSql & vbcrlf & ",X2Rst11"
				strSql = strSql & vbcrlf & ",X2Rst12"
				strSql = strSql & vbcrlf & ",X2Rst01"
				strSql = strSql & vbcrlf & ",X2Rst02"
				strSql = strSql & vbcrlf & ",X2Rst03"
				strSql = strSql & vbcrlf & ",(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",X2Pln04"
				strSql = strSql & vbcrlf & ",X2Pln05"
				strSql = strSql & vbcrlf & ",X2Pln06"
				strSql = strSql & vbcrlf & ",X2Pln07"
				strSql = strSql & vbcrlf & ",X2Pln08"
				strSql = strSql & vbcrlf & ",X2Pln09"
				strSql = strSql & vbcrlf & ",X2Pln10"
				strSql = strSql & vbcrlf & ",X2Pln11"
				strSql = strSql & vbcrlf & ",X2Pln12"
				strSql = strSql & vbcrlf & ",X2Pln01"
				strSql = strSql & vbcrlf & ",X2Pln02"
				strSql = strSql & vbcrlf & ",X2Pln03"
				strSql = strSql & vbcrlf & ",(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK�ʏ�Ǘ���"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C2Rst04"
				strSql = strSql & vbcrlf & ",C2Rst05"
				strSql = strSql & vbcrlf & ",C2Rst06"
				strSql = strSql & vbcrlf & ",C2Rst07"
				strSql = strSql & vbcrlf & ",C2Rst08"
				strSql = strSql & vbcrlf & ",C2Rst09"
				strSql = strSql & vbcrlf & ",C2Rst10"
				strSql = strSql & vbcrlf & ",C2Rst11"
				strSql = strSql & vbcrlf & ",C2Rst12"
				strSql = strSql & vbcrlf & ",C2Rst01"
				strSql = strSql & vbcrlf & ",C2Rst02"
				strSql = strSql & vbcrlf & ",C2Rst03"
				strSql = strSql & vbcrlf & ",(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C2Pln04"
				strSql = strSql & vbcrlf & ",C2Pln05"
				strSql = strSql & vbcrlf & ",C2Pln06"
				strSql = strSql & vbcrlf & ",C2Pln07"
				strSql = strSql & vbcrlf & ",C2Pln08"
				strSql = strSql & vbcrlf & ",C2Pln09"
				strSql = strSql & vbcrlf & ",C2Pln10"
				strSql = strSql & vbcrlf & ",C2Pln11"
				strSql = strSql & vbcrlf & ",C2Pln12"
				strSql = strSql & vbcrlf & ",C2Pln01"
				strSql = strSql & vbcrlf & ",C2Pln02"
				strSql = strSql & vbcrlf & ",C2Pln03"
				strSql = strSql & vbcrlf & ",(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "-(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK���ʊǗ���"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C9Rst04"
				strSql = strSql & vbcrlf & ",C9Rst05"
				strSql = strSql & vbcrlf & ",C9Rst06"
				strSql = strSql & vbcrlf & ",C9Rst07"
				strSql = strSql & vbcrlf & ",C9Rst08"
				strSql = strSql & vbcrlf & ",C9Rst09"
				strSql = strSql & vbcrlf & ",C9Rst10"
				strSql = strSql & vbcrlf & ",C9Rst11"
				strSql = strSql & vbcrlf & ",C9Rst12"
				strSql = strSql & vbcrlf & ",C9Rst01"
				strSql = strSql & vbcrlf & ",C9Rst02"
				strSql = strSql & vbcrlf & ",C9Rst03"
				strSql = strSql & vbcrlf & ",(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C9Pln04"
				strSql = strSql & vbcrlf & ",C9Pln05"
				strSql = strSql & vbcrlf & ",C9Pln06"
				strSql = strSql & vbcrlf & ",C9Pln07"
				strSql = strSql & vbcrlf & ",C9Pln08"
				strSql = strSql & vbcrlf & ",C9Pln09"
				strSql = strSql & vbcrlf & ",C9Pln10"
				strSql = strSql & vbcrlf & ",C9Pln11"
				strSql = strSql & vbcrlf & ",C9Pln12"
				strSql = strSql & vbcrlf & ",C9Pln01"
				strSql = strSql & vbcrlf & ",C9Pln02"
				strSql = strSql & vbcrlf & ",C9Pln03"
				strSql = strSql & vbcrlf & ",(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "-(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK�V�X�e����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",DRst04"
				strSql = strSql & vbcrlf & ",DRst05"
				strSql = strSql & vbcrlf & ",DRst06"
				strSql = strSql & vbcrlf & ",DRst07"
				strSql = strSql & vbcrlf & ",DRst08"
				strSql = strSql & vbcrlf & ",DRst09"
				strSql = strSql & vbcrlf & ",DRst10"
				strSql = strSql & vbcrlf & ",DRst11"
				strSql = strSql & vbcrlf & ",DRst12"
				strSql = strSql & vbcrlf & ",DRst01"
				strSql = strSql & vbcrlf & ",DRst02"
				strSql = strSql & vbcrlf & ",DRst03"
				strSql = strSql & vbcrlf & ",(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",DPln04"
				strSql = strSql & vbcrlf & ",DPln05"
				strSql = strSql & vbcrlf & ",DPln06"
				strSql = strSql & vbcrlf & ",DPln07"
				strSql = strSql & vbcrlf & ",DPln08"
				strSql = strSql & vbcrlf & ",DPln09"
				strSql = strSql & vbcrlf & ",DPln10"
				strSql = strSql & vbcrlf & ",DPln11"
				strSql = strSql & vbcrlf & ",DPln12"
				strSql = strSql & vbcrlf & ",DPln01"
				strSql = strSql & vbcrlf & ",DPln02"
				strSql = strSql & vbcrlf & ",DPln03"
				strSql = strSql & vbcrlf & ",(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03)"
				strSql = strSql & vbcrlf & "-(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK�Œ��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "+(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "+(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "+(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",X2Rst04+C2Rst04+C9Rst04+DRst04"
				strSql = strSql & vbcrlf & ",X2Rst05+C2Rst05+C9Rst05+DRst05"
				strSql = strSql & vbcrlf & ",X2Rst06+C2Rst06+C9Rst06+DRst06"
				strSql = strSql & vbcrlf & ",X2Rst07+C2Rst07+C9Rst07+DRst07"
				strSql = strSql & vbcrlf & ",X2Rst08+C2Rst08+C9Rst08+DRst08"
				strSql = strSql & vbcrlf & ",X2Rst09+C2Rst09+C9Rst09+DRst09"
				strSql = strSql & vbcrlf & ",X2Rst10+C2Rst10+C9Rst10+DRst10"
				strSql = strSql & vbcrlf & ",X2Rst11+C2Rst11+C9Rst11+DRst11"
				strSql = strSql & vbcrlf & ",X2Rst12+C2Rst12+C9Rst12+DRst12"
				strSql = strSql & vbcrlf & ",X2Rst01+C2Rst01+C9Rst01+DRst01"
				strSql = strSql & vbcrlf & ",X2Rst02+C2Rst02+C9Rst02+DRst02"
				strSql = strSql & vbcrlf & ",X2Rst03+C2Rst03+C9Rst03+DRst03"
				strSql = strSql & vbcrlf & ",(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)"
				strSql = strSql & vbcrlf & "+(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03)"
				strSql = strSql & vbcrlf & "+(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03)"
				strSql = strSql & vbcrlf & "+(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",X2Pln04+C2Pln04+C9Pln04+DPln04"
				strSql = strSql & vbcrlf & ",X2Pln05+C2Pln05+C9Pln05+DPln05"
				strSql = strSql & vbcrlf & ",X2Pln06+C2Pln06+C9Pln06+DPln06"
				strSql = strSql & vbcrlf & ",X2Pln07+C2Pln07+C9Pln07+DPln07"
				strSql = strSql & vbcrlf & ",X2Pln08+C2Pln08+C9Pln08+DPln08"
				strSql = strSql & vbcrlf & ",X2Pln09+C2Pln09+C9Pln09+DPln09"
				strSql = strSql & vbcrlf & ",X2Pln10+C2Pln10+C9Pln10+DPln10"
				strSql = strSql & vbcrlf & ",X2Pln11+C2Pln11+C9Pln11+DPln11"
				strSql = strSql & vbcrlf & ",X2Pln12+C2Pln12+C9Pln12+DPln12"
				strSql = strSql & vbcrlf & ",X2Pln01+C2Pln01+C9Pln01+DPln01"
				strSql = strSql & vbcrlf & ",X2Pln02+C2Pln02+C9Pln02+DPln02"
				strSql = strSql & vbcrlf & ",X2Pln03+C2Pln03+C9Pln03+DPln03"
				strSql = strSql & vbcrlf & " ((X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "+(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "+(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "+(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03))"
				strSql = strSql & vbcrlf & "-((X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)"
				strSql = strSql & vbcrlf & "+(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03)"
				strSql = strSql & vbcrlf & "+(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03)"
				strSql = strSql & vbcrlf & "+(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK�c�Ɨ��v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum((if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "YK�Ζ�����","YK�Ζ�����(��)","YK�Ζ�����(�v)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(DT<='" & strYM & "',Result,Plan)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(DT<='" & strYM & "',Result,Plan)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(Prev",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan","","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(Plan",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan","","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(DT<='" & strYM & "',Result,Plan)-Plan) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				select case strKubun
				case "YK�Ζ�����"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM101')"
				case "YK�Ζ�����(��)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM102')"
				case "YK�Ζ�����(�v)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD in ('TM101','TM102'))"
				end select
			case "YK��Ǝ���","YK��Ǝ���(��)","YK��Ǝ���(�v)"
				dim	strTM101
				dim	strTM201
				select case strKubun
				case "YK��Ǝ���"
					strTM101 = "'TM101'"
					strTM201 = "'TM201'"
				case "YK��Ǝ���(��)"
					strTM101 = "'TM102'"
					strTM201 = "'TM202'"
				case "YK��Ǝ���(�v)"
					strTM101 = "'TM101','TM102'"
					strTM201 = "'TM201','TM202'"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM201 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK���Ǝ���","YK���Ǝ���(��)","YK���Ǝ���(�v)"
				dim	strTM301
				select case strKubun
				case "YK���Ǝ���"
					strTM101 = "'TM101'"
					strTM301 = "'TM301'"
				case "YK���Ǝ���(��)"
					strTM101 = "'TM102'"
					strTM301 = "'TM302'"
				case "YK���Ǝ���(�v)"
					strTM101 = "'TM101','TM102'"
					strTM301 = "'TM301','TM301'"
				end select
	
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM301 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK�L������","YK�L������(��)","YK�L������(�v)"
				dim	strTM401
				select case strKubun
				case "YK�L������"
					strTM101 = "'TM101'"
					strTM401 = "'TM401'"
				case "YK�L������(��)"
					strTM101 = "'TM102'"
					strTM401 = "'TM402'"
				case "YK�L������(�v)"
					strTM101 = "'TM101','TM102'"
					strTM401 = "'TM401','TM402'"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM401 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK����H��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD <> 'Y9999'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD <> 'Y9999'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "YK�H��(�]�T����)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD <> 'Y9999'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD <> 'Y9999'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			end select
		elseif left(strKubun,2) = "CK" then
			select case strKubun
			case "CK����"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'A%'"
			case "CK���㌴��"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'B%'"
			case "CK���ڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
			case "CK����"
				strSql = "select" 
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "CK���E���v"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "CK�Ԑڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'X0200'"
			case "CK�ʏ�Ǘ���"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD in ('C0500','C0600')"
			case "CK���ʊǗ���"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'C9999'"
			case "CK�V�X�e����"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'D0100'"
			case "CK�Œ��"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "CK�c�Ɨ��v"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "CK�Ζ�����"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM101'"
			case "CK�Ζ�����(��)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM102'"
			case "CK��Ǝ���"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM201'"
			case "CK��Ǝ���(��)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM202'"
			case "CK���Ǝ���"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM301'"
			case "CK���Ǝ���(��)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM302'"
			case "CK�L������"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM401'"
			case "CK�L������(��)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM402'"
			case "CK����H��"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'Y9999'"
			case "CK�H��(�]�T����)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'Y9999'"
			end select
		else
			select case strKubun
			case "JK����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'A%'"
			case "����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","A%","")
			case "JK���㌴��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "       )"
			case "JK���ڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result        * if(KamokuCD = 'X0200',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan          * if(KamokuCD = 'X0200',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result * if(KamokuCD = 'X0200',-1,1),0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result        * if(KamokuCD = 'X0200',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan          * if(KamokuCD = 'X0200',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result * if(KamokuCD = 'X0200',-1,1),0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK���E���v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan   * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK�Ԑڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "       )"
			case "JK�ʏ�Ǘ���"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "       )"
			case "JK���ʊǗ���"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "       )"
			case "JK�V�X�e����"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "       )"
			case "JK�Œ��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "JK�c�Ɨ��v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan   * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "JK�Ζ�����","JK�Ζ�����(��)","JK�Ζ�����(�v)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				select case strKubun
				case "JK�Ζ�����"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM101'"
				case "JK�Ζ�����(��)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM102'"
				case "JK�Ζ�����(�v)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM100'"
				end select
				strSql = strSql & vbcrlf & "       )"
	'			strSql = "select"
	'			strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result,0)) A_Result"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
	'			strSql = strSql & vbcrlf & ",sum(Result) T_Result"
	'			strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
	'			strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
	'			strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
	'			strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
	'			strSql = strSql & vbcrlf & " from IrData"
	'			strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
	'			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
	'			strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y0010'"
	'			strSql = strSql & vbcrlf & "       )"
			case "JK��Ǝ���","JK��Ǝ���(��)"
	'			dim	strTM101
	'			dim	strTM201
				select case strKubun
				case "JK��Ǝ���"
					strTM101 = "TM101"
					strTM201 = "TM201"
				case "JK��Ǝ���(��)"
					strTM101 = "TM102"
					strTM201 = "TM202"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM201 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK���Ǝ���","JK���Ǝ���(��)"
	'			dim	strTM301
				select case strKubun
				case "JK���Ǝ���"
					strTM101 = "TM101"
					strTM301 = "TM301"
				case "JK���Ǝ���(��)"
					strTM101 = "TM102"
					strTM301 = "TM302"
				end select
	
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM301 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK�L������","JK�L������(��)"
	'			dim	strTM401
				select case strKubun
				case "JK�L������"
					strTM101 = "TM101"
					strTM401 = "TM401"
				case "JK�L������(��)"
					strTM101 = "TM102"
					strTM401 = "TM402"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM401 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK���ԓ���"
				strSql = "select"
				strSql = strSql & vbcrlf & " k.KamokuName"
				strSql = strSql & vbcrlf & ",PersonCD"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(0) A_Result_Hi_0"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(0) A_Plan_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) A_Margin_0"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(0) A_Prev_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) A_Prev_Margin_0"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(0) T_Result_Hi_0"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(0) T_Plan_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) T_Margin_0"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(0) T_Prev_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) T_Prev_Margin_0"
				strSql = strSql & vbcrlf & " from Attendance a"
				strSql = strSql & vbcrlf & " inner join Kamoku k on (a.KamokuCD = k.KamokuCD)"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (a.KamokuCD like 'TM%1' or a.KamokuCD like 'TM%2')"

				strSqlAdd = strSqlAdd & vbcrlf & " group by"
				strSqlAdd = strSqlAdd & vbcrlf & " a.KamokuCD"
				strSqlAdd = strSqlAdd & vbcrlf & ",k.KamokuName"
				strSqlAdd = strSqlAdd & vbcrlf & ",a.PersonCD"
				strSqlAdd = strSqlAdd & vbcrlf & " having sum(Result) <> 0 or sum(Prev) <> 0 or sum(Plan) <> 0"
				strSqlAdd = strSqlAdd & vbcrlf & " order by"
				strSqlAdd = strSqlAdd & vbcrlf & " a.KamokuCD"
				strSqlAdd = strSqlAdd & vbcrlf & ",a.PersonCD"
			case "JK����H��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "JK�H��(�]�T����)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "���ޔ�"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B0100,B0200,B0500","")
			case "�H���d��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B0300","")
			case "���̑��d��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B%","")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","-B0100,B0200,B0300,B0500","")
			case "���ڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C - a.Prev_X2) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C - a.Plan_X2) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C - a.Result_X2) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C - a.tPrev_X2) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C - a.tPlan_X2) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C - a.tResult_X2) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "����"	'���v
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_B + (a.Prev_C - a.Prev_X2)) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_B + (a.Plan_C - a.Plan_X2)) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_B + (a.Result_C - a.Result_X2)) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_B + (a.tPrev_C - a.tPrev_X2)) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_B + (a.tPlan_C - a.tPlan_X2)) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_B + (a.tResult_C - a.tResult_X2)) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "���E���v"	'����|����
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_A - a.Prev_B - (a.Prev_C - a.Prev_X2)) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_A - a.Plan_B - (a.Plan_C - a.Plan_X2)) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_A - a.Result_B - (a.Result_C - a.Result_X2)) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_A - a.tPrev_B - (a.tPrev_C - a.tPrev_X2)) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_A - a.tPlan_B - (a.tPlan_C - a.tPlan_X2)) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_A - a.tResult_B - (a.tResult_C - a.tResult_X2)) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "�Ԑڐl����"
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_X2) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_X2) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_X2) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_X2) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_X2) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_X2) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "�ʏ�Ǘ���"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","C0500,C0600","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C5 + a.Prev_C6) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C5 + a.Plan_C6) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C5 + a.Result_C6) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C5 + a.tPrev_C6) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C5 + a.tPlan_C6) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C5 + a.tResult_C6) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "���ʊǗ���"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","C9999","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C9) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C9) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C9) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C9) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C9) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C9) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "�V�X�e����"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","D0100","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_D1) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_D1) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_D1) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_D1) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_D1) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_D1) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "�Œ��"	'���v
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_X2 + a.Prev_C5 + a.Prev_C9 + a.Prev_D1) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_X2 + a.Plan_C5 + a.Plan_C9 + a.Plan_D1) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_X2 + a.Result_C5 + a.Result_C9 + a.Result_D1) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_X2 + a.tPrev_C5 + a.tPrev_C9 + a.tPrev_D1) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_X2 + a.tPlan_C5 + a.tPlan_C9 + a.tPlan_D1) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_X2 + a.tResult_C5 + a.tResult_C9 + a.tResult_D1) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "�o��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'C%'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "�d��"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B%","")
			case "�e���v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev * if(KamokuCD like 'B%',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'B%',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'B%',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD like 'B%',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD like 'B%',-1,1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD like 'B%',-1,1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD like 'B%',-1,1)) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD like 'B%',-1,1)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     OR KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     )"
			case "�c�Ɨ��v"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","A0000","D9999")
			case else
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","99999","")
			end select
		end if
		if right(strSql,3) <> ") i" then
			if strJKubun <> "" then
				strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
			else
				strSql = strSql & vbcrlf & "   AND SyushiCd <> ''"
			end if
			strSql = strSql & SqlWhere("and", "SyushiCD", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		end if
	end if
	MakeSql = strSql & strSqlAdd
End Function
'-------------------------------------------------------------
'1�N�Ԃ�select ��field����Ԃ�
'-------------------------------------------------------------
Function GetYearMonthYK(byVal strFld,byVal strSum,byval strAnd,byval strMul,byVal strYM,byVal strY)
	select case GetRequest("ptype","")
	case "pTableJKYearMonth"
'		if strSum = "Plan" then
'			strYM = clng(strYM) + 100
'		end if
	case "pTableJKYearMonth2"
	case else
		GetYearMonthYK = ""
		exit function
	end select
	dim	strYearMonthYK
	strYearMonthYK = ""
	dim	i
	for	i = 1 to 12
		dim	m
		m = i + 3
		if m > 12 then
			m = m - 12
		end if
		dim	strYYYYMM
		strYYYYMM = GetNendo(strYM,m)
		dim	strF
		strF = strSum
		if strSum = "Result,Plan" then
			if CLng(strYYYYMM) <= CLng(strYM) then
				strF = split(strSum,",")(0)
			else
				strF = split(strSum,",")(1)
			end if
		end if
		dim	strSelect
		strSelect = ",sum(if(" & strFld & " = '" & strYYYYMM & "'" & strAnd & "," & strF & ",0)" & strMul & ")"
		strYearMonthYK = strYearMonthYK & strSelect & vbCrLf
	next
	GetYearMonthYK = strYearMonthYK
End Function

'-------------------------------------------------------------
'1�N�Ԃ�select ��field����Ԃ�
'-------------------------------------------------------------
Function GetYearMonthSF(byVal strFld,byVal strSum,byVal strYM,byVal iOffset)
	dim	strYearMonthSF
	strYearMonthSF = ""
	select case GetRequest("ptype","")
	case "pTableJKYearMonth","pTableJKYearMonth2"
		dim	lngYYYYMM
		lngYYYYMM = CLng(strYM)
		lngYYYYMM = lngYYYYMM + (iOffset * 100)
	
		dim	strY
		if inStr(strSum,"Prev") > 0 then
			strY = "x"
		else
			strY = "y"
		end if
		dim	i
		for	i = 1 to 12
			dim	m
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strYYYYMM
			strYYYYMM = GetNendo(CStr(lngYYYYMM),m)
			strYearMonthSF = strYearMonthSF & strSum & " * (if(" & strFld & "='" & strYYYYMM & "',1,0))) " & strY & strYYYYMM & vbCrLf
		next
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		dim	iKeikaku
		iKeikaku = 0
		dim	strM
		if right(GetRequest("ptype",""),2) = "11" then
			iKeikaku = 8
		else
			strM = right(GetRequest("ptype",""),1)
			if strM <> "u" then
				iKeikaku = CLng(strM)
			end if
			if iKeikaku > 4 then
				' 4��=1 4-3
				' 5��=2 5-3
				' 6��=3 6-3
				' 7��=4 7-3
				' 8��=5 8-3
				' 9��=6 9-3
				'10��=7 10-3
				'11��=8 11-3
				'12��=9 12-3
				' 1��=10 1+9
				' 2��=11 2+9
				' 3��=12 3+9
				iKeikaku = iKeikaku - 3
			else
				iKeikaku = iKeikaku + 9
			end if
		end if
		dim	strKeikakuYM
		strKeikakuYM = ""
		dim	strComma
		strComma = " "
		for	i = 1 to 12
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			strYYYYMM = GetNendo(strYM,m)
			if i >= iKeikaku then
				if strKeikakuYM = "" then
					strKeikakuYM = strYYYYMM
				end if
				strSum = Replace(strSum,"Result","Plan")
			end if
			strYearMonthSF = strYearMonthSF & strComma & "sum( Round(if(" & strFld & " = '" & strYYYYMM & "'," & strSum & ",0),0) ) A_Result_" & m & vbCrLf
			strComma = ","
		next
		if strSum = "Plan" then
			strSum = ""
		else
			strSum = Replace(strSum,"Plan","")
		end if
		strYearMonthSF = strYearMonthSF & ",sum( Round(if(" & strFld & " < '" & strKeikakuYM & "',Result,Plan)" & strSum & ",0) ) T_Result" & vbCrLf
	end select
	GetYearMonthSF = strYearMonthSF
End Function

'-------------------------------------------------------------
'1�N�Ԃ�<TH>�^�O��Ԃ�
'-------------------------------------------------------------
Function GetYearMonthTH(byVal strYM,byVal iOffset)
	dim	strYearMonthTH
	strYearMonthTH = ""
	select case GetRequest("ptype","")
	case "pTableJKYearMonth","pTableJKYearMonth2"
		dim	lngYYYYMM
		lngYYYYMM = CLng(strYM)
		lngYYYYMM = lngYYYYMM + (iOffset * 100)
	
		dim	i
		for	i = 1 to 12
			dim	m
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strYYYYMM
			strYYYYMM = GetNendo(CStr(lngYYYYMM),m)
			strYearMonthTH = strYearMonthTH & "<TH>" & strYYYYMM & "</TH>" & vbCrLf
		next
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		dim	iKeikaku
		iKeikaku = 0
		if right(GetRequest("ptype",""),2) = "11" then
			iKeikaku = 8
		else
			strM = right(GetRequest("ptype",""),1)
			if strM <> "u" then
				iKeikaku = CLng(strM)
			end if
			if iKeikaku > 4 then
				' 4��=1 4-3
				' 5��=2 5-3
				' 6��=3 6-3
				' 7��=4 7-3
				' 8��=5 8-3
				' 9��=6 9-3
				'10��=7 10-3
				'11��=8 11-3
				'12��=9 12-3
				' 1��=10 1+9
				' 2��=11 2+9
				' 3��=12 3+9
				iKeikaku = iKeikaku - 3
			else
				iKeikaku = iKeikaku + 9
			end if
		end if
		for	i = 1 to 12
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strM
			strM = m & "��"
			if i >= iKeikaku then
				strM = strM & "<br>�v��"
			end if
			strYearMonthTH = strYearMonthTH & "<TH>" & strM & "</TH>" & vbCrLf
		next
	end select
	GetYearMonthTH = strYearMonthTH
End Function

'-------------------------------------------------------------
'�e�[�u���w�b�_�[
'-------------------------------------------------------------
Function MakeHeader(byVal objDb,byVal strCenterCD,byVal strYM,byval strTableType)
	dim	strHeader
	dim	objRs
	dim	strSyushiCD
	dim	strSyushiName
	dim	i
	strHeader = vbCrLf
	strHeader = strHeader & "<!-- MakeHear(" & strCenterCD & "," & strYM & "," & strTableType & ")-->" & vbCrLf

	select case strTableType
	case "pList"
		strHeader = strHeader & GetTdList(objDb,strCenterCD,"Header",strYM,strTableType)
	case "pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"
		dim	strPeriod1
		dim	strPeriod2
		dim	intYM1
		dim	intYM2
		if strTableType = "pTableJKYearMonth2" then
			intYM1 = 0	' �������ʂ�
			intYM2 = 0	' �����v��
		else
			intYM1 = 0	' �������ʂ�
			intYM2 = 1	' �����v��
		end if
		strPeriod1 = GetPeriod(strYM) + intYM1
		strPeriod2 = GetPeriod(strYM) + intYM2

		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strYM & "</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strPeriod1 & "�����ʂ�</TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,intYM1)
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strPeriod2 & "���v��</TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,intYM2)
		strHeader = strHeader & "<TH colspan=""1"" rowspan=""1"">��</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTable"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & "<TH>" & GetPeriod(strYM) - 1 & "��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""3"" title=""" & strYM & """>" & GetPeriod(strYM) & "��</TH>" & vbCrLf
		strHeader = strHeader & "<TH>" & GetPeriod(strYM) - 1 & "���݌v</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""3"">" & GetPeriod(strYM) & "���݌v</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH>����</TH>" & vbCrLf
		strHeader = strHeader & "<TH>�v��</TH>" & vbCrLf
		strHeader = strHeader & "<TH>����</TH>" & vbCrLf
		strHeader = strHeader & "<TH>��</TH>" & vbCrLf
		strHeader = strHeader & "<TH>����</TH>" & vbCrLf
		strHeader = strHeader & "<TH>�v��</TH>" & vbCrLf
		strHeader = strHeader & "<TH>����</TH>" & vbCrLf
		strHeader = strHeader & "<TH>��</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTableJK","pTableJKKan"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">��������</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">���ƌv��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">�O�N����</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">�݌v����</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">���ƌv��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">�O�N����</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">��</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,0)
'		strHeader = strHeader & "<TH colspan=""1"">4��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">5��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">6��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">7��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">8��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">9��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">10��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">11��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">12��<br>�v��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">1��<br>�v��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">2��<br>�v��</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">3��<br>�v��</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">���v</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	end select

	strHeader = strHeader & "<!-- MakeHear() End -->" & vbCrLf

	MakeHeader = strHeader
End Function
'-------------------------------------------------------------
'�G���[���b�Z�[�WHTML
'-------------------------------------------------------------
Function MakeError(byVal objErr)
	dim	strHTML
	strHTML = strHTML & "<tr><td>Err.Number:</td>"
	strHTML = strHTML & "<td>0x" & Hex(objErr.Number) & "(" & objErr.Number & ")</td></tr>"
	strHTML = strHTML & "<tr><td>Err.Description:</td>"
	strHTML = strHTML & "<td>" & objErr.Description & "</td></tr>"
	strHTML = strHTML & "<tr><td>Err.Source:</td>"
	strHTML = strHTML & "<td>" & objErr.Source & "</td></tr>"
	MakeError = strHTML
End Function
'-------------------------------------------------------------
'���x�}�X�^�[����SQL
'-------------------------------------------------------------
Function SyushiSql(byVal strCenterCD)
	dim	strSql

	strSql = "select"
	strSql = strSql & " s.CenterCD CenterCD"
	strSql = strSql & ",s.SyushiKB"
	strSql = strSql & ",sk.SyushiKBName"
	strSql = strSql & ",s.SyushiCD SyushiCD"
	strSql = strSql & ",s.SyushiName SyushiName"
	strSql = strSql & " FROM Syushi s"
	strSql = strSql & " left outer join SyushiKB sk on (s.SyushiKB = sk.SyushiKB)"
	strSql = strSql & " WHERE s.CenterCD = '" & strCenterCD &"'"
	strSql = strSql & " ORDER BY"
	strSql = strSql & " s.CenterCD"
	strSql = strSql & ",s.SyushiKB"
	strSql = strSql & ",s.SyushiCD"

	SyushiSql = strSql
End Function
'-------------------------------------------------------------
'���n�p select �t�B�[���h
'sum(if(YM <= '201211',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Total
'-------------------------------------------------------------
Function ChakuSelect(byVal strSql,byVal strSum1,byVal strSum2,byVal strFld)
	dim	i
	dim	sYM
	dim	strV
	for i = 1 to 12
		strSql = strSql & vbcrlf
		if i = 1 then
			strSql = strSql & " "
		else
			strSql = strSql & ","
		end if
		if i < 10 then
			sYM = "2012" & right("0" & i + 3,2)
		else
			sYM = "2013" & right("0" & i - 9,2)
		end if
		if sYM <= "201211" then
			strV = "Result"
		else
			strV = "Plan"
		end if
		strV = "if(" & strFld & " = '" & sYM & "'," & strV & ",0)"
		strSql = strSql & strSum1 & strV & strSum2 & " A_" & sYM
	next
	strSql = strSql & ","
	strV = "if(" & strFld & " <= '201211',Result,Plan)"
	strSql = strSql & strSum1 & strV & strSum2 & " A_Total"
	ChakuSelect = strSql
End Function
'-------------------------------------------------------------
'�N��
'-------------------------------------------------------------
Function GetNendo(byVal strYM,byVal intM)
	dim	intYear
	dim	intMonth
	intYear = CInt(left(strYM,4))
	intMonth = CInt(right(strYM,2))
	if intMonth < 4 then
		intYear = intYear - 1
	end if
	if intM < 4 then
		intYear = intYear + 1
	end if
	GetNendo = intYear & right("0" & intM,2)
End Function
%>
