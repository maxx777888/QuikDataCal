#pragma once
#include <OpenXLSX.hpp>
#include <xlnt/xlnt.hpp>
#include <iostream>
#include <cmath>
#include <numeric>
#include <iomanip>
#include <sstream>
#include <vector>
#include <string>
#include <algorithm>
#include "stock.h"



class fileXLSX
{
public:

	fileXLSX(const std::string& fileName_, const std::string& sheetName_);
	~fileXLSX();
	std::vector<StockDataClass> getAllDataFromFile();
	void printStartData();//�������� ������ ������ �� �������
	//�������� ����������� ������ ������������ �� ���. ����� � �������� �����. ���� �������������� �� ����� �����
	void printCalculatedData(double number, double riskProc); 
	//�������� ����������� ������ ����� ������ ������������ �� ���. ����� � �������� �����. ���� �������������� �� ����� �����
	void printShareCalData(const std::string& comName, double amount, double riskProc);

	//����������� ����������� ��� ��� ����� � ������� � ������ 
	void calculateOptLotSize(const std::string& comName, double amount, double riskProc);

	//����������� ����������� ��� ��� ����� � ������� �� ������ � �������
	void calculateOptLotSizeForAllTable(double amount, double riskProc);

	void makeDealBook(std::vector<std::string>& v, double amount, double riskProc); //������� ����� Deal � �������
	
private:
	std::string fileName = "";
	std::string sheetName = "";
	////////////////////// �������� �������� � �������
	std::string sharePriceNameStr = "���� ������.";
	std::string shareCodeNameStr = "��� �����������";
	std::string shareLotNameStr = "���";
	std::string priceMoveNameStr = "��� ����";
	std::string bestSupplyNameStr = "����. ����";
	std::string bestDemandNameStr = "����. �����";
	std::string minPriceStr = "���. ����";
	std::string maxPriceStr = "����. ����";
	std::string assetVolumeStr = "������";
	std::string openPriceStr = "���� ����.";
	/////////////////////////////////////////////
	void calData(double number, double riskProc);
	void printShare(StockCalculateDataClass& data);
	//////////////////////////////////////////////////
	xlnt::rgb_color yellow_color = xlnt::rgb_color(255, 253, 79);
	xlnt::rgb_color red_color = xlnt::rgb_color(255, 102, 102);
	xlnt::rgb_color perpule_color = xlnt::rgb_color(153, 0, 255);
	xlnt::rgb_color white_color = xlnt::rgb_color(255, 255, 255);
	xlnt::rgb_color blue_color = xlnt::rgb_color(0, 0, 139);
	//////////////////////////////////////////////////////////
	xlnt::border cell_border;
	xlnt::border::border_property cell_border_property;

	xlnt::alignment cell_alignment_center;
	

	OpenXLSX::XLDocument openDoc;
	xlnt::workbook createdDoc;

	std::vector<StockDataClass> stockVec;
	std::vector<StockCalculateDataClass> stockCalculatedVec;
	std::vector<StockDataOptDialStopProfitClass> stockDataOptDialProfitVec;

	int getColumnNumber(const std::string& columnName);//�� �������� ����� ������� ����� �������
	int countFractionalDigits(double number); //������������ ���-�� ���� ����� ������� � ������� �����
	
	//������� �� ����� ���������� �� ������
	void printDealData(StockDataOptDialStopProfitClass data);

	bool createCalculatedTableFromTableData(double amount, double riskProc);

	void makeNoFormulaCellStyle(xlnt::cell& c);//������ ����� ������ ��� �������
	void makeCellStyle(xlnt::cell& c);//������ ����� ������ � ��������

	bool findComNameInShareVec(const std::string com_name);//���������� ������ ���� ������� ��� �������� � ������� �������
};

