# Prediction of Call and Put Option Prices on Bitcoin and S&P 500 Using Machine Learning and Monte Carlo Simulation  

## 📌 Introduction  
This repository contains the code used in my thesis, which allowed me to earn my Mathematics degree and graduate on time (3.5 years).  

In this study, I compare the prediction of **European call and put option prices** for two different assets—**Bitcoin and S&P 500**—using two **machine learning models**:  
- **eXtreme Gradient Boosting (XGBoost)**  
- **Light Gradient Boosting Machine (LightGBM)**  

Additionally, I employ **Monte Carlo simulation** to estimate option prices based on **stochastic asset price modeling**.  

## 📊 Machine Learning Approach  
For machine learning, the independent variables (**X**) consist of:  
- **Three volatility estimation methods**:  
  - **Bipower Variation (BV)**  
  - **Realized Variance (RV)**  
  - **Signed Jumps Variation (SJV)**  
- **Four different lag periods**: **1, 7, 15, and 30 days**  

### 🔧 Hyperparameter Tuning  
To optimize model performance, I use two methods:  
1. **DataRobot** (automated ML platform)  
2. **GridSearchCV** (manual hyperparameter tuning)  

If the results differ between the two approaches, both models are tested on the **testing dataset** to determine which performs better. The dataset is shuffled and split into an **80:20 ratio (training:testing)**.  

### 📈 Feature Analysis  
To understand the **impact of volatility estimation and lag periods**, I analyze models where only **one type of volatility measure** or **one specific lag** is used. This helps identify which volatility measure and lag have the most influence on option prices.  

## 🎲 Monte Carlo Simulation Approach  
Unlike machine learning, **Monte Carlo simulation** requires **time-ordered data** and does not allow random shuffling.  

### 🔢 Steps:  
1. **Modeling log-returns** of the underlying assets (**Bitcoin futures and S&P 500**) using the **Generalized Hyperbolic (GH) distribution**.  
2. **Simulating asset price paths** using stochastic processes.  
3. **Generating 10,000 simulations** to compute the **percentile range (5% to 95%)** for option prices (95% confidence interval).  
4. **Verifying whether real option prices** fall within this confidence interval.  

### 📌 Data Partitioning Strategies:  
- **5 days before maturity**  
- **80:20 data split**  
- **60:40 data split**  

## 🏆 Results  
✅ **XGBoost outperforms LightGBM**, achieving lower **RMSE and MAPE values** across most datasets.  
✅ **Short-term lags** are **more relevant for Bitcoin**, as its **high volatility** is best captured by shorter intervals.  
✅ **Long-term lags** are **better for S&P 500**, due to its relatively stable market movement.  
✅ **Monte Carlo's 60:40 data split** provides the **most accurate confidence range** for option prices.  
✅ **Bitcoin option prices** are **more predictable** than those of the S&P 500.  
✅ **Call options** exhibit **better predictive accuracy** than **put options**.  

## 🚀 Conclusion  
These findings provide valuable insights for **option issuers and investors**, helping them choose the **most effective prediction methods** for different assets. By combining **machine learning and Monte Carlo simulations**, this study supports **better investment decision-making** for Bitcoin and S&P 500 options.  

