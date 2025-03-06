# Load libraries
library(readxl)
library(dplyr)
library(ghyp)
library(stats)
library(writexl)

# Read data
#call_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Call")
#put_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Put")
asset_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Asset Price R")
test_call_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx', sheet = "Testing Data Call")
test_put_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx', sheet = "Testing Data Put")

# Ubah kolom tanggal menjadi format Date
asset_bitcoin$Time <- as.Date(asset_bitcoin$Time, format = "%Y-%m-%d")
test_call_btc$Time <- as.Date(test_call_btc$Time, format = "%Y-%m-%d")

# Menghitung log return dari underlying asset Bitcoin
returns_btc <- diff(log(asset_bitcoin$Last))
returns_btc <- data.frame(Date = asset_bitcoin$Time[-1], Return = returns_btc)

# Membagi data training (60%) dan testing (40%) berdasarkan urutan
n <- nrow(returns_btc)
train_size <- floor(0.6 * n)
train_returns <- returns_btc[1:train_size, ]  # Data training
test_returns <- returns_btc[(train_size + 1):n, ]  # Data testing

# Fit Generalized Hyperbolic distribution pada data training
fit_ghyp_train <- fit.ghypuv(train_returns$Return, silent = FALSE)

# Set parameter simulasi
num_simulations <- 10000
num_observations <- nrow(test_returns)

# Simulasi Monte Carlo
simulated_returns <- matrix(NA, nrow = num_simulations, ncol = num_observations)
for (i in 1:num_simulations) {
  simulated_returns[i, ] <- rghyp(num_observations, fit_ghyp_train)
}

# Konversi simulated_returns ke data frame
simulated_returns_df <- as.data.frame(simulated_returns)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_6040.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Simulated Returns" = simulated_returns_df), path = output_path)

# Konfirmasi penyimpanan selesai
cat("Simulated returns telah disimpan ke:", output_path, "\n")

# Path ke file Excel yang berisi simulated_returns
input_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_6040.xlsx"

# Membaca data simulated_returns dari Excel
simulated_returns <- read_excel(input_path, sheet = "Simulated Returns")

# Konversi simulated_returns ke matrix untuk digunakan dalam perhitungan
simulated_returns <- as.matrix(simulated_returns)

# Inisialisasi harga awal dan time steps
S0_call <- 39195 #Harga aset di tgl 2023-11-21 (baris terakhir training data)
price_paths <- matrix(NA, nrow = num_simulations, ncol = num_observations+1)
price_paths[, 1] <- S0_call

# Menghitung price path berdasarkan hasil simulasi
for (sim in 1:num_simulations) {
  for (day in 2:(num_observations+1)) {
    price_paths[sim, day] <- price_paths[sim, day - 1] * exp(simulated_returns[sim, day - 1])
  }
}

# Konversi pricepaths ke data frame
price_paths_df <- as.data.frame(price_paths)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Pricepath_BTC_6040.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Pricepath" = price_paths_df), path = output_path)

# Plot untuk beberapa jalur harga simulasi hingga 20 kolom pertama
plot(1:151, price_paths[1, 1:151], type = "l", col = "blue", ylim = range(price_paths[, 1:151]), 
     xlab = "Hari", ylab = "Harga Bitcoin", main = "Simulasi Jalur Harga Bitcoin (151 Hari Terakhir)")
# Tambahkan jalur harga lain ke plot
for (sim in 2:10000) {  # Misalnya, plotkan 200 jalur simulasi pertama
  lines(1:151, price_paths[sim, 1:151], col = sample(colors(), 1))
}

# Menghitung rata-rata price path
average_path <- apply(price_paths, 2, mean)

# Membuat data frame untuk plot
comparison_prices <- data.frame(
  Date = test_returns$Date,
  Actual_Price = asset_bitcoin$Last[(train_size):n],
  Average_Simulated_Price = average_path
)

# Plot harga aktual dan rata-rata hasil simulasi price paths
library(ggplot2)
# Definisikan tanggal untuk plot sesuai dengan data testing
tanggal_aset <- asset_bitcoin$Time
tanggal_simulasi <- tanggal_aset[(train_size):n]  # Rentang tanggal untuk data simulasi

# Plot harga aktual aset Bitcoin
plot(tanggal_aset, asset_bitcoin$Last, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga Bitcoin",
     main = "Harga Aset Bitcoin dan Rata-rata Simulasi GH vs Normal")

# Tambahkan garis untuk hasil simulasi GH dan normal pada data testing
lines(tanggal_simulasi, average_path, col = "red", lwd = 2)

# Tambahkan legenda
legend("topleft", legend = c("Harga Aset Bitcoin", "Simulasi GH"),
       col = c("blue", "red"), lwd = 2)


##### RMSE MAPE #####
# Menghitung harga opsi Call untuk setiap baris di data testing
r <- 0.0461  # Suku bunga tahunan
call_prices <- numeric(nrow(test_call_btc))  # Inisialisasi vektor untuk menyimpan harga opsi

for (i in 1:nrow(test_call_btc)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_call_btc$Strike[i]
  days_to_maturity <- test_call_btc$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations+1]
  payoff_call <- pmax(0, ST - K)
  
  # Menghitung harga opsi dengan diskon faktor
  call_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_call)
}

# Cetak hasil harga opsi untuk setiap baris di data testing
print(call_prices)

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual
results_btc_call <- data.frame(
  Time = test_call_btc$Time,
  Actual_Call_Price = test_call_btc$Last,
  Simulated_Call_Price = call_prices
)

# Cetak hasil perbandingan
print(results_btc_call)

# Hitung RMSE
rmse_btc_call <- sqrt(mean((results_btc_call$Actual_Call_Price - results_btc_call$Simulated_Call_Price)^2))
cat("RMSE:", rmse_btc_call, "\n")

# Hitung MAPE
mape_btc_call <- mean(abs((results_btc_call$Actual_Call_Price - results_btc_call$Simulated_Call_Price) / results_btc_call$Actual_Call_Price)) * 100
cat("MAPE:", mape_btc_call, "%\n")

# Menghitung harga opsi Put untuk setiap baris di data testing
r <- 0.0461  # Suku bunga tahunan
put_prices <- numeric(nrow(test_put_btc))  # Inisialisasi vektor untuk menyimpan harga opsi Put

for (i in 1:nrow(test_put_btc)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_put_btc$Strike[i]
  days_to_maturity <- test_put_btc$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations+1]
  payoff_put <- pmax(0, K - ST)
  
  # Menghitung harga opsi Put dengan diskon faktor
  put_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_put)
}

# Cetak hasil harga opsi Put untuk setiap baris di data testing
print(put_prices)

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual
results_put <- data.frame(
  Time = test_put_btc$Time,
  Actual_Put_Price = test_put_btc$Last,
  Simulated_Put_Price = put_prices
)

# Cetak hasil perbandingan
print(results_put)

# Hitung RMSE
rmse_put <- sqrt(mean((results_put$Actual_Put_Price - results_put$Simulated_Put_Price)^2))
cat("RMSE:", rmse_put, "\n")

# Hitung MAPE
mape_put <- mean(abs((results_put$Actual_Put_Price - results_put$Simulated_Put_Price) / results_put$Actual_Put_Price)) * 100
cat("MAPE:", mape_put, "%\n")

# Tentukan path file untuk menyimpan hasil
library(writexl)
file_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results MC/60:40/Hasil MC6040 BTC.xlsx"

# Menyimpan kedua data frame dalam satu file Excel dengan dua sheet
write_xlsx(list(
  "Put Option Results" = results_put,
  "Call Option Results" = results_btc_call
), path = file_path)


#### CI 95% ####
r <- 0.0461
# Misalkan jumlah simulasi adalah 10.000
num_simulations <- nrow(price_paths)

# Matriks untuk menyimpan harga opsi Call langsung dari payoff
call_prices_matrix <- matrix(0, nrow = num_simulations, ncol = nrow(test_call_btc))

for (i in 1:nrow(test_call_btc)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_call_btc$Strike[i]
  days_to_maturity <- test_call_btc$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations + 1]  # Harga akhir simulasi
  payoff_call <- pmax(0, ST - K)  # Payoff Call = max(0, ST - K)
  
  # Menghitung harga opsi Call untuk setiap simulasi (bukan rata-rata)
  call_prices_matrix[, i] <- exp(-r * days_to_maturity) * payoff_call
}

# Output: call_prices_matrix berisi harga opsi Call untuk setiap simulasi dan setiap baris data

# Path untuk menyimpan file Excel
output_call_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_Price_BTC_6040.xlsx"

# Simpan call_prices_matrix ke Excel
write_xlsx(list("Call Prices" = as.data.frame(call_prices_matrix)), path = output_call_prices_path)
cat("call_prices_matrix telah disimpan ke:", output_call_prices_path, "\n")

# ---------------------------------------
# Proses membagi data per hari dan menghitung rata-rata
# ---------------------------------------

# Path file Excel yang berisi call_prices_matrix
input_call_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_Price_BTC_6040.xlsx"

# Membaca data dari sheet "Call Prices" dan menyimpannya sebagai call_prices_matrix
call_prices_matrix <- as.matrix(read_excel(input_call_prices_path, sheet = "Call Prices"))

# Tentukan jumlah hari berdasarkan kolom di call_prices_matrix
num_days <- ncol(call_prices_matrix)

# Tentukan jumlah data dalam satu kelompok
group_size <- 100

# Inisialisasi data frame untuk menyimpan rata-rata
average_prices_per_group <- data.frame(Day = integer(), Group = integer(), Average = numeric())

# Loop untuk menghitung rata-rata setiap kelompok per hari
for (day in 1:num_days) {
  # Ambil data untuk hari tersebut
  day_data <- call_prices_matrix[, day]
  
  # Tentukan jumlah grup pada hari tersebut
  num_groups <- floor(length(day_data) / group_size)
  
  for (group in 1:num_groups) {
    # Ambil data untuk grup tertentu
    group_data <- day_data[((group - 1) * group_size + 1):(group * group_size)]
    
    # Hitung rata-rata grup
    group_average <- mean(group_data)
    
    # Tambahkan ke data frame hasil
    average_prices_per_group <- rbind(
      average_prices_per_group,
      data.frame(Day = day, Group = group, Average = group_average)
    )
  }
}

# Path untuk menyimpan file Excel hasil rata-rata
output_average_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Average_Prices_BTC_6040.xlsx"

# Simpan rata-rata ke file Excel
write_xlsx(list("Average Prices" = average_prices_per_group), path = output_average_prices_path)
cat("Rata-rata harga per grup telah disimpan ke:", output_average_prices_path, "\n")

# ---------------------------------------
# Inisialisasi data frame untuk hasil uji normalitas
# ---------------------------------------

normality_results <- data.frame(Day = integer(), PValue = numeric(), Normal = logical())

# Loop untuk melakukan uji KS setiap hari
unique_days <- unique(average_prices_per_group$Day)

for (day in unique_days) {
  # Ambil data rata-rata untuk hari tertentu
  day_data <- average_prices_per_group$Average[average_prices_per_group$Day == day]
  
  # Uji normalitas menggunakan Kolmogorov-Smirnov
  ks_test <- ks.test(day_data, "pnorm", mean(day_data), sd(day_data))
  
  # Simpan hasil uji ke data frame
  normality_results <- rbind(
    normality_results,
    data.frame(
      Day = day,
      PValue = ks_test$p.value,
      Normal = ks_test$p.value > 0.05 # Normal jika p-value > 0.05
    )
  )
}

# ---------------------------------------
# Simpan hasil uji normalitas ke Excel
# ---------------------------------------

# Path untuk menyimpan file Excel hasil uji normalitas
output_normality_results_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Normality_Results_BTC_6040.xlsx"

# Simpan hasil uji normalitas ke file Excel
write_xlsx(list("Normality Results" = normality_results), path = output_normality_results_path)
cat("Hasil uji normalitas telah disimpan ke:", output_normality_results_path, "\n")

# Inisialisasi data frame untuk menyimpan hasil rata-rata dan confidence interval per hari
confidence_intervals_per_day <- data.frame(Day = integer(), 
                                           Average = numeric(), 
                                           CI_Lower = numeric(), 
                                           CI_Upper = numeric())

# Loop untuk menghitung rata-rata dan confidence interval per hari dari average_prices_per_group
for (day in unique(average_prices_per_group$Day)) {
  # Ambil seluruh data rata-rata untuk hari tersebut
  day_data <- average_prices_per_group$Average[average_prices_per_group$Day == day]
  #day_data <- average_prices_per_group$Average[average_prices_per_group$Day == 1]
  
  # Hitung rata-rata dan standar deviasi dari seluruh data untuk hari tersebut
  day_average <- mean(day_data)
  day_sd <- sd(day_data)
  day_n <- length(day_data)
  
  # Hitung batas bawah dan atas dari confidence interval 95% untuk hari tersebut
  ci_lower <- day_average - 1.96 * (day_sd / sqrt(day_n))
  ci_upper <- day_average + 1.96 * (day_sd / sqrt(day_n))
  
  # Tambahkan ke data frame hasil
  confidence_intervals_per_day <- rbind(
    confidence_intervals_per_day,
    data.frame(Day = day, Average = day_average, 
               CI_Lower = ci_lower, CI_Upper = ci_upper)
  )
}

# Tambahkan kolom "Call Price" dari test_call_btc$Last ke confidence_intervals_per_day
confidence_intervals_per_day$Call_Price <- test_call_btc$Last

# Tambahkan kolom untuk mengecek apakah Call Price berada di rentang CI
confidence_intervals_per_day$Within_CI <- with(
  confidence_intervals_per_day,
  Call_Price >= ci_lower & Call_Price <= ci_upper
)

# Tambahkan kolom "Asset Price" dari test_call_btc$Asset_Price ke confidence_intervals_per_day
confidence_intervals_per_day$Asset_Price <- test_call_btc$`Asset Price`

# Tambahkan kolom untuk mengecek apakah Asset Price berada di rentang CI
confidence_intervals_per_day$Within_CI_Asset <- with(
  confidence_intervals_per_day,
  Asset_Price >= ci_lower & Asset_Price <= ci_upper
)

# Periksa hasil
head(confidence_intervals_per_day)

# Simpan hasil confidence interval ke file Excel
output_ci_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Confidence_Intervals_Per_Day_BTC_6040.xlsx"

# Simpan data ke Excel
write_xlsx(list("Confidence Intervals" = confidence_intervals_per_day), path = output_ci_path)
cat("Confidence intervals per day telah disimpan ke:", output_ci_path, "\n")

## PUT ##
# Menghitung harga opsi Put menggunakan simulasi Monte Carlo
r <- 0.0461
num_simulations <- nrow(price_paths)

put_prices_matrix <- matrix(0, nrow = num_simulations, ncol = nrow(test_put_btc))

for (i in 1:nrow(test_put_btc)) {
  K <- test_put_btc$Strike[i]
  days_to_maturity <- test_put_btc$Maturity[i]
  
  ST <- price_paths[, num_observations + 1]  # Harga akhir simulasi
  payoff_put <- pmax(0, K - ST)  # Payoff Put = max(0, K - ST)
  
  put_prices_matrix[, i] <- exp(-r * days_to_maturity) * payoff_put
}

output_put_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_PricePUT_BTC_6040.xlsx"
write_xlsx(list("Put Prices" = as.data.frame(put_prices_matrix)), path = output_put_prices_path)
cat("put_prices_matrix telah disimpan ke:", output_put_prices_path, "\n")

# Membaca data dan menghitung rata-rata harga per grup
input_put_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_PricePUT_BTC_6040.xlsx"
put_prices_matrix <- as.matrix(read_excel(input_put_prices_path, sheet = "Put Prices"))

num_days <- ncol(put_prices_matrix)
group_size <- 100

average_prices_per_group <- data.frame(Day = integer(), Group = integer(), Average = numeric())

for (day in 1:num_days) {
  day_data <- put_prices_matrix[, day]
  num_groups <- floor(length(day_data) / group_size)
  
  for (group in 1:num_groups) {
    group_data <- day_data[((group - 1) * group_size + 1):(group * group_size)]
    group_average <- mean(group_data)
    average_prices_per_group <- rbind(
      average_prices_per_group,
      data.frame(Day = day, Group = group, Average = group_average)
    )
  }
}

output_average_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Average_PricesPUT_BTC_6040.xlsx"
write_xlsx(list("Average Prices" = average_prices_per_group), path = output_average_prices_path)
cat("Rata-rata harga per grup telah disimpan ke:", output_average_prices_path, "\n")

# Uji normalitas menggunakan Kolmogorov-Smirnov
normality_results <- data.frame(Day = integer(), PValue = numeric(), Normal = logical())

# Mendapatkan daftar hari yang unik
unique_days <- unique(average_prices_per_group$Day)

# Loop untuk melakukan uji normalitas setiap hari
for (day in unique_days) {
  # Ambil data rata-rata untuk hari tertentu
  day_data <- average_prices_per_group$Average[average_prices_per_group$Day == day]
  
  # Pastikan data tidak mengandung NA dan memiliki ukuran yang cukup besar
  day_data <- na.omit(day_data)  # Menghapus NA
  if (length(day_data) < 2) {  # Uji KS memerlukan lebih dari 1 data point
    next  # Jika data terlalu sedikit, lewati hari ini
  }
  
  # Lakukan uji normalitas dengan Kolmogorov-Smirnov
  ks_test <- ks.test(day_data, "pnorm", mean(day_data), sd(day_data))
  
  # Simpan hasil uji ke data frame
  normality_results <- rbind(
    normality_results,
    data.frame(
      Day = day,
      PValue = ks_test$p.value,
      Normal = ks_test$p.value > 0.05  # Normal jika p-value > 0.05
    )
  )
}

# Initialize the data frame for storing results
normality_results <- data.frame(Day = integer(), PValue = numeric(), Normal = logical())

# Loop through each unique day to perform normality tests
for (day in unique_days) {
  day_data <- average_prices_per_group$Average[average_prices_per_group$Day == day]
  day_data <- na.omit(day_data)  # Remove NA values
  
  # Check if all values are identical
  if (length(unique(day_data)) == 1) {
    # If all values are identical, flag as "Non-normal" (no variability)
    normality_results <- rbind(
      normality_results,
      data.frame(
        Day = day,
        PValue = NA,  # No valid p-value due to no variability
        Normal = FALSE
      )
    )
  } else {
    # If data has variability, perform the Shapiro-Wilk test
    shapiro_test <- shapiro.test(day_data)
    
    normality_results <- rbind(
      normality_results,
      data.frame(
        Day = day,
        PValue = shapiro_test$p.value,
        Normal = shapiro_test$p.value > 0.05  # Normal if p-value > 0.05
      )
    )
  }
}

# View the results
print(normality_results)

# Simpan hasil uji normalitas ke Excel
output_normality_results_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Normality_Results_PUT_BTC_6040.xlsx"
write_xlsx(list("Normality Results" = normality_results), path = output_normality_results_path)
cat("Hasil uji normalitas PUT telah disimpan ke:", output_normality_results_path, "\n")


# Menghitung confidence intervals per hari
confidence_intervals_per_day <- data.frame(Day = integer(), 
                                           Average = numeric(), 
                                           CI_Lower = numeric(), 
                                           CI_Upper = numeric())

for (day in unique(average_prices_per_group$Day)) {
  day_data <- average_prices_per_group$Average[average_prices_per_group$Day == day]
  
  day_average <- mean(day_data)
  day_sd <- sd(day_data)
  day_n <- length(day_data)
  
  ci_lower <- day_average - 1.96 * (day_sd / sqrt(day_n))
  ci_upper <- day_average + 1.96 * (day_sd / sqrt(day_n))
  
  confidence_intervals_per_day <- rbind(
    confidence_intervals_per_day,
    data.frame(Day = day, Average = day_average, 
               CI_Lower = ci_lower, CI_Upper = ci_upper)
  )
}

# Tambahkan kolom "Put Price" dari test_put_btc$Last ke confidence_intervals_per_day
confidence_intervals_per_day$Put_Price <- test_put_btc$Last

# Tambahkan kolom untuk mengecek apakah Put Price berada di rentang CI
confidence_intervals_per_day$Within_CI <- with(
  confidence_intervals_per_day,
  Put_Price >= ci_lower & Put_Price <= ci_upper
)

# Tambahkan kolom "Asset Price" dari test_put_btc$`Asset Price` ke confidence_intervals_per_day
confidence_intervals_per_day$Asset_Price <- test_put_btc$`Asset Price`

# Tambahkan kolom untuk mengecek apakah Asset Price berada di rentang CI
confidence_intervals_per_day$Within_CI_Asset <- with(
  confidence_intervals_per_day,
  Asset_Price >= ci_lower & Asset_Price <= ci_upper
)

# Periksa hasil
head(confidence_intervals_per_day)

# Simpan hasil confidence interval ke file Excel
output_ci_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Confidence_Intervals_PUT_Per_Day_BTC_6040.xlsx"
write_xlsx(list("Confidence Intervals" = confidence_intervals_per_day), path = output_ci_path)
cat("Confidence intervals per day telah disimpan ke:", output_ci_path, "\n")


#### Menggunakan Percentile ####
# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo
price_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Kumpulan_Price_BTC_6040.xlsx")

# 2. Membaca dataset Testing Data Call
test_call_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx', sheet = "Testing Data Call")

# 3. Menghitung quantile 5% dan 95% untuk setiap kolom di price_data
quantiles <- apply(price_data, 2, function(column) {
  quantile(column, probs = c(0.05, 0.95), na.rm = TRUE)
})

# Konversi hasil quantile menjadi data frame
quantile_df <- as.data.frame(t(quantiles))
colnames(quantile_df) <- c("Percentile_5%", "Percentile_95%")

# Tambahkan kolom quantile ke data price_data
# price_data <- cbind(price_data, quantile_df)

# 4. Tambahkan kolom `Actual_Price` ke quantile_df dari test_call_btc$Last
# Pastikan jumlah kolom di quantile_df sesuai dengan jumlah kolom di price_data
if (nrow(quantile_df) == nrow(test_call_btc)) {
  quantile_df$Actual_Price <- test_call_btc$Last
} else {
  stop("Jumlah kolom di quantile_df dan test_call_btc tidak sesuai. Harap cek data.")
}

# 5. Tambahkan kolom pengecekan apakah `Actual_Price` berada dalam rentang quantile
quantile_df$InRange <- quantile_df$Actual_Price >= quantile_df$'Percentile_5%' & 
  quantile_df$Actual_Price <= quantile_df$'Percentile_95%'

# 6. Simpan hasil ke file Excel baru
write.xlsx(quantile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Percentile_BTCCall_6040.xlsx")


# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo untuk Put
priceput_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Kumpulan_PricePUT_BTC_6040.xlsx")

# 2. Membaca dataset Testing Data Put
test_put_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx', sheet = "Testing Data Put")

# 3. Menghitung quantile 5% dan 95% untuk setiap kolom di priceput_data
quantiles_put <- apply(priceput_data, 2, function(column) {
  quantile(column, probs = c(0.05, 0.95), na.rm = TRUE)
})

# Konversi hasil quantile menjadi data frame
quantileput_df <- as.data.frame(t(quantiles_put))
colnames(quantileput_df) <- c("Percentile_5%", "Percentile_95%")

# 4. Tambahkan kolom `Actual_Price` ke quantileput_df dari test_put_sp500$Last
# Pastikan jumlah kolom di quantileput_df sesuai dengan jumlah kolom di priceput_data
if (nrow(quantileput_df) == nrow(test_put_btc)) {
  quantileput_df$Actual_Price <- test_put_btc$Last
} else {
  stop("Jumlah kolom di quantileput_df dan test_put_btc tidak sesuai. Harap cek data.")
}

# 5. Tambahkan kolom pengecekan apakah `Actual_Price` berada dalam rentang quantile
quantileput_df$InRange <- quantileput_df$Actual_Price >= quantileput_df$'Percentile_5%' & 
  quantileput_df$Actual_Price <= quantileput_df$'Percentile_95%'

# 6. Simpan hasil ke file Excel baru
write.xlsx(quantileput_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Percentile_BTCPut_6040.xlsx")

##### Cek percentile untuk harga aset #####
library(readxl)
library(dplyr)
library(ggplot2)

# 1. Baca data simulasi
pricepath_data <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Pricepath_BTC_6040.xlsx')

# 2. Baca data actual
asset_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Asset Price R")

# 3. Bagi data menjadi 60% training dan 40% testing
n <- nrow(asset_bitcoin)
train_size <- round(0.6 * n)
train_data <- asset_bitcoin[1:train_size, ]
test_data <- asset_bitcoin[(train_size + 1):n, ]

# 4. Gabungkan baris terakhir training dan seluruh testing sebagai actual data
actual_data <- rbind(train_data[train_size, ], test_data)

# 5. Hitung percentile 5% dan 95% untuk setiap kolom di data simulasi
percentile_df <- data.frame(
  Column = names(pricepath_data),
  Percentile_5 = apply(pricepath_data, 2, function(x) quantile(x, 0.05, na.rm = TRUE)),
  Percentile_95 = apply(pricepath_data, 2, function(x) quantile(x, 0.95, na.rm = TRUE))
)

# 6. Bandingkan nilai actual dengan percentile
actual_prices <- actual_data$Last # Sesuaikan kolom actual
percentile_df$Actual_Price <- actual_prices

percentile_df$In_Range <- mapply(
  function(actual, p5, p95) actual >= p5 & actual <= p95,
  actual_prices,
  percentile_df$Percentile_5,
  percentile_df$Percentile_95
)

# 7. Plot percentiles dan harga aktual
tanggal_aset <- actual_data$Time  # Sesuaikan dengan nama kolom tanggal
tanggal_simulasi <- tanggal_aset[(train_size):n]  # Rentang tanggal untuk data testing

ggplot(data = NULL, aes(x = tanggal_aset)) +
  geom_line(aes(y = actual_data$Last, color = "Harga Aset BTC"), size = 1) +
  geom_line(aes(y = percentile_df$Percentile_5, color = "Percentile 5%"), linetype = "dashed") +
  geom_line(aes(y = percentile_df$Percentile_95, color = "Percentile 95%"), linetype = "dashed") +
  labs(
    title = "Harga Aset BTC (Testing 40%) dan Percentile 5%-95%",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  theme_minimal()

ggplot() +
  # Plot harga aset untuk training data
  geom_line(data = train_data, aes(x = Time, y = Last, color = "Training Data"), size = 1) +
  # Plot harga aset untuk testing data
  geom_line(data = test_data, aes(x = Time, y = Last, color = "Testing Data"), size = 1) +
  
  # Label dan tampilan
  labs(
    title = "Harga Aset BTC Rasio 60:40 (Training vs Testing)",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  scale_color_manual(values = c("Training Data" = "blue", "Testing Data" = "red")) +
  theme_minimal()

# 8. Simpan data hasil dengan percentile dan status
write.xlsx(percentile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC6040/Percentile_Result_AsetBTC_6040.xlsx")
