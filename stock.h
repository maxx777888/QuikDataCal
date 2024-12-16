#pragma once
#include <iostream>
#include <string>
#include <iomanip>
#include <sstream>

struct Stock {
    std::string companyCode = ""; // Название компании
    double sharePrice = 0.f;       // Цена одной акции
    int lot = 0; //Кол-во акций в одном лоте
    double lotPrice = 0.f; //Цена в деньгах одного лота
    double priceMove = 0.f; //Шаг цены
    double lotMove = 0.f; //Шаг цены в деньгах одного лота
    double supplyPrice = 0.f;
    double demandPrice = 0.f;
    double minPrice = 0.f;
    double maxPrice = 0.f;
    double volume = 0.f;
    double openPrice = 0.f;
    //double forecastingDailyPriceRanges = 0.f;//Прогнозирование дневных диапазонов цен
    //double forecastingMaxPriceForTomorrow = 0.f;
    //double forecastingMinPriceForTomorrow = 0.f;
};

struct StockDil {
    std::string companyCode = "";//Код компании
    double currentPrice = 0.f;
    int oneLot = 0;
    double accNum = 0.f;//Суммы для сделки
    int lotToBuy = 0;//Кол-во лотов, которое максимум можно купить на сумму сделки
    double procToRist = 0.f;//максимальный процент риска от суммы сделки
    double riskAmount = 0.f;//максимальная сумма риска, процент риска от суммы сделки
    int priceMoveToStop = 0;//Кол-во шагов которое может пройти цена пока не достигнет суммы максимального процента риска
    double stopPrice = 0.f;//Цена стоп лоса
};

struct StockOptDialStopProfit {
    std::string companyCode = "";//Код компании
    double investAmount = 0.f;//Инвестиционная сумма
    double riskProcent = 0.f; //Процент риска нужен для инвестиционной суммы
    double companyPrice = 0.f; //Текущая цена компании
    double maxLotInDeal = 0.f;//Максимальное кол-во лотов которое можно купить для формирования сделки
    double maxRiskInDeal = 0.f;//Максимально возможные потери
    double maxUseOfMoneyInDeal = 0.f;// Максимальное использование инветиционной суммы в сделке
    double freeMoneyFromDeal = 0.f;//Остаток средств от инветиционной суммы
    double maxProfitInDeal = 0.f;//Максимальный профит
    double riskProfitRatio = 0.f;//Коэффициент Риск/Профит, если меньше 1, то прибыль больше, если больше то риск больше
};

class StockDataOptDialStopProfitClass {
public:
    StockDataOptDialStopProfitClass(const std::string& comName_, double amount_, double riskProc_, double price_,
        double maxLotsInDeal_, double maxRiskInDeal_, double maxUseOfMoneyInDeal_,
        double freeMoneyFromDeal_, double maxProfitInDeal_, double riskProfitRatio_) {

        stockData.companyCode = comName_;
        stockData.investAmount = amount_;
        stockData.riskProcent = riskProc_;
        stockData.companyPrice = price_;
        stockData.maxLotInDeal = maxLotsInDeal_;
        stockData.maxRiskInDeal = maxRiskInDeal_;
        stockData.maxUseOfMoneyInDeal = maxUseOfMoneyInDeal_;
        stockData.freeMoneyFromDeal = freeMoneyFromDeal_;
        stockData.maxProfitInDeal = maxProfitInDeal_;
        stockData.riskProfitRatio = riskProfitRatio_;
    }

    void printStockDataOptDialStopProfit();

    int countFractionalDigits(double number); //Подсчитывает кол-во цифр после запятой в дробном числе

    std::string getCompanyName() const { return stockData.companyCode; }
    double getInvestAmount() const { return stockData.investAmount; }
    double getRiskProcent() const { return stockData.riskProcent; }
    double getCompanyPrice() const { return stockData.companyPrice; }
    double getMaxLotInDeal() const { return stockData.maxLotInDeal; }
    double getMaxRiskInDeal() const { return stockData.maxRiskInDeal; }
    double getMaxUseOfMoneyInDeal() const { return stockData.maxUseOfMoneyInDeal; }
    double getFreeMoneyFromDeal() const { return stockData.freeMoneyFromDeal; }
    double getMaxProfitInDeal() const { return stockData.maxProfitInDeal; }
    double getRiskProfitRatio() const { return stockData.riskProfitRatio; }

private:
    StockOptDialStopProfit stockData;
};

class StockCalculateDataClass {
public:
    StockCalculateDataClass(std::string companyCode_, double accNum_, int lot_, double lotPrice_, double procToRist_,
        double lotMove_, double sharePrice_, double priceMove_)
    {
        if (sharePrice_ == 0) {
            dialStockData.companyCode = companyCode_;
            dialStockData.currentPrice = 0;
            dialStockData.oneLot = lot_;
            dialStockData.accNum = 0;
            dialStockData.lotToBuy = 0;
            dialStockData.procToRist = 0;
            dialStockData.riskAmount = 0;
            dialStockData.priceMoveToStop = 0;
            dialStockData.stopPrice = 0;
        }
        else {
            dialStockData.companyCode = companyCode_;
            dialStockData.currentPrice = sharePrice_;
            dialStockData.accNum = accNum_;
            dialStockData.oneLot = lot_;
            dialStockData.lotToBuy = static_cast<int>(accNum_ / lotPrice_);
            dialStockData.procToRist = procToRist_ / 100;
            dialStockData.riskAmount = accNum_ * dialStockData.procToRist;
            dialStockData.priceMoveToStop = static_cast<int>(dialStockData.riskAmount / (lotMove_ * dialStockData.lotToBuy));
            dialStockData.stopPrice = sharePrice_ - (dialStockData.priceMoveToStop * priceMove_);
        }
        
    }
    void printStockCalculatedData() const;
    void printOneStockCalData() const;
private:
    StockDil dialStockData;
};


class StockDataClass {
public:
    StockDataClass(std::string companyCode_, double sharePrice_, int lot_, double priceMove_,
        double supplyPrice_, double demandPrice_, double minPrice_, double maxPrice_, double volume_, 
        double openPrice_) {

        stockInfo.companyCode = companyCode_;
        stockInfo.sharePrice = sharePrice_;
        stockInfo.lot = lot_;
        stockInfo.lotPrice = sharePrice_ * static_cast<double>(lot_);
        stockInfo.priceMove = priceMove_;
        stockInfo.lotMove = priceMove_ * static_cast<double>(lot_);
        stockInfo.supplyPrice = supplyPrice_;
        stockInfo.demandPrice = demandPrice_;
        stockInfo.minPrice = minPrice_;
        stockInfo.maxPrice = maxPrice_;
        stockInfo.volume = volume_;
        stockInfo.openPrice = openPrice_;

        /*int fDPR = 0;
        if (sharePrice_ < openPrice_) { fDPR = 1; } 
        else if (sharePrice_ > openPrice_)
        { fDPR = 2;}
        else if (sharePrice_ == openPrice_) { fDPR = 3; }

        if (fDPR == 1)
        {
            stockInfo.forecastingDailyPriceRanges = (maxPrice_ + minPrice_ + sharePrice_) / 2;
            stockInfo.forecastingMaxPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - minPrice_;
            stockInfo.forecastingMinPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - maxPrice_;
        }
        else if (fDPR == 2)
        {
            stockInfo.forecastingDailyPriceRanges = (maxPrice_ + minPrice_ + sharePrice_ + maxPrice_) / 2;
            stockInfo.forecastingMaxPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - minPrice_;
            stockInfo.forecastingMinPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - maxPrice_;
        }
        else if (fDPR == 3)
        {
            stockInfo.forecastingDailyPriceRanges = (maxPrice_ + minPrice_ + sharePrice_ + sharePrice_) / 2;
            stockInfo.forecastingMaxPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - minPrice_;
            stockInfo.forecastingMinPriceForTomorrow = stockInfo.forecastingDailyPriceRanges - maxPrice_;
        }*/

        
    }

    std::string getCompanyName() const { return stockInfo.companyCode; }
    void setCompanyName(const std::string& companyName) { stockInfo.companyCode = companyName; }

    double getSharePrice() const { return stockInfo.sharePrice; }
    
    double getInOneLotShares() const { return stockInfo.lot; }
    double getLotPrice() const { return stockInfo.lotPrice; }
    double getLotMove() const { return stockInfo.lotMove; }
    double getPriceMove() const { return stockInfo.priceMove; }
    double getSupplyPrice() const { return stockInfo.supplyPrice; }
    double getDemandPrice() const { return stockInfo.demandPrice; }
    double getOpenPrice() const { return stockInfo.openPrice; }
    double getMaxPrice() const { return stockInfo.maxPrice; }
    double getMinPrice() const { return stockInfo.minPrice; }
    double getVolume() const { return stockInfo.volume; }

    void setSharePrice(double sharePrice) { stockInfo.sharePrice = sharePrice; }
    void printStockData() const;
    

private:
    Stock stockInfo;
};