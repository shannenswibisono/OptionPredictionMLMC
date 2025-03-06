# Load libraries
library(readxl)
library(dplyr)
library(ghyp)
library(stats)
library(ggplot2)
library(writexl)

# Read data
test_call_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx', sheet = "Testing Data")
test_put_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx', sheet = "Testing Data Put")
asset_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500_Data.xlsx', sheet="Asset Price All")

# Ubah kolom tanggal menjadi format Date
asset_sp500$Date <- as.Date(asset_sp500$Date, format = "%Y-%m-%d")
test_call_sp500$Time <- as.Date(test_call_sp500$Time, format = "%Y-%m-%d")
test_put_sp500$Time <- as.Date(test_put_sp500$Time, format = "%Y-%m-%d")

########### BARU #########
# Menghitung log return dari harga penutupan
returns_sp500 <- diff(log(asset_sp500$GSPC.Close)) 
returns_sp500 <- data.frame(Date = asset_sp500$Date[-1], Return = returns_sp500)

# Pisahkan data training (sebelum 15 Juli 2024) dan data testing (15-19 Juli 2024)
train_returns <- returns_sp500 %>%
  filter(Date < as.Date("2024-07-15"))
test_returns <- returns_sp500 %>%
  filter(Date >= as.Date("2024-07-15") & Date <= as.Date("2024-07-19"))

# Pastikan data memiliki variasi yang cukup
if (nrow(train_returns) == 0 || nrow(test_returns) == 0) {
  stop("Data training atau testing kosong!")
}
if (sd(train_returns$Return) == 0) {
  stop("Variabilitas data training terlalu kecil, fitting tidak dapat dilakukan.")
}

fitgh = fit.ghypuv(train_returns$Return)
# Fit distribusi Normal pada data training
# install.packages("fitdistrplus")
library(fitdistrplus)
fit_norm_train <- fitdist(train_returns$Return, "norm")
cat("Parameter Distribusi Normal:\n")
print(fit_norm_train)

# Periksa parameter hasil fitting
cat("Parameter distribusi GH:\n")
print(fitgh)

# Simulasi Monte Carlo
num_simulations <- 10000
num_observations <- nrow(test_returns)

# Mengambil parameter hasil fitting distribusi Normal
fit_mean <- fit_norm_train$estimate["mean"]
fit_sd <- fit_norm_train$estimate["sd"]

# Simulasi return menggunakan distribusi Normal
simulated_returns <- matrix(NA, nrow = num_simulations, ncol = num_observations)
for (i in 1:num_simulations) {
  simulated_returns[i, ] <- rnorm(num_observations, mean = fit_mean, sd = fit_sd)
}

# Konversi simulated_returns ke data frame
simulated_returns_df <- as.data.frame(simulated_returns)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Simulated_Return_SP500_5DTM.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Simulated Returns" = simulated_returns_df), path = output_path)

# Konfirmasi penyimpanan selesai
cat("Simulated returns telah disimpan ke:", output_path, "\n")

# Path ke file Excel yang berisi simulated_returns
input_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Simulated_Return_SP500_5DTM.xlsx"

# Membaca data simulated_returns dari Excel
simulated_returns <- read_excel(input_path, sheet = "Simulated Returns")

# Konversi simulated_returns ke matrix untuk digunakan dalam perhitungan
simulated_returns <- as.matrix(simulated_returns)

# Inisialisasi harga awal
S0_call <- 5615.350098 #Harga akhir training data (2024-07-12)

# Jalur harga berdasarkan hasil simulasi
price_paths <- matrix(NA, nrow = num_simulations, ncol = num_observations+1)
price_paths[, 1] <- S0_call

# Menghitung price path
for (sim in 1:num_simulations) {
  for (day in 2:(num_observations+1)) {
    price_paths[sim, day] <- price_paths[sim, day - 1] * exp(simulated_returns[sim, day - 1])
  }
}

# Konversi pricepaths ke data frame
price_paths_df <- as.data.frame(price_paths)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Pricepath_SP500_5DTM.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Pricepath" = price_paths_df), path = output_path)

# Plot untuk beberapa jalur harga simulasi hingga 20 kolom pertama
plot(1:6, price_paths[1, 1:6], type = "l", col = "blue", ylim = range(price_paths[, 1:6]), 
     xlab = "Hari", ylab = "Harga S&P500", main = "Simulasi Jalur Harga S&P500 (6 Hari Terakhir)")
# Tambahkan jalur harga lain ke plot
for (sim in 2:1000) {  # Misalnya, plotkan 200 jalur simulasi pertama
  lines(1:6, price_paths[sim, 1:6], col = sample(colors(), 1))
}

# Menghitung rata-rata price path
average_path <- apply(price_paths, 2, mean)

# Tentukan tanggal untuk plot sesuai dengan data testing
tanggal_aset <- asset_sp500$Date

# Tentukan tanggal batas training dan testing
train_date <- as.Date("2024-07-15")
train_size <- which(tanggal_aset == train_date) - 1  # Ukuran training set (jumlah data sebelum 15 Juli 2024)
n <- length(tanggal_aset)  # Jumlah total data

# Pisahkan tanggal simulasi (untuk testing)
tanggal_simulasi <- tanggal_aset[(train_size):n]

# Membuat data frame untuk plot
# Menggabungkan harga aktual dan harga simulasi, dengan memasukkan nilai NA pada bagian training untuk Average_Simulated_Price
comparison_prices <- data.frame(
  Date = tanggal_aset,
  Actual_Price = asset_sp500$GSPC.Close,
  Average_Simulated_Price = c(rep(NA, train_size-1), average_path)  # Isi NA untuk data sebelum simulasi dimulai
)

# Plot harga aktual dan rata-rata hasil simulasi price paths
library(ggplot2)
# Definisikan tanggal untuk plot sesuai dengan data testing
tanggal_aset <- asset_sp500$Date
tanggal_simulasi <- tanggal_aset[(train_size):n]  # Rentang tanggal untuk data testing

# Plot harga aktual aset Bitcoin
# Plot harga aktual aset S&P500
plot(tanggal_aset, asset_sp500$GSPC.Close, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga S&P500", 
     main = "Harga Aset S&P500 dan Rata-rata Simulasi Normal", 
     xaxt = "n")  # Menghilangkan grid dan menyesuaikan sumbu x

# Tambahkan garis untuk hasil simulasi Normal
lines(tanggal_simulasi, average_path, col = "orange", lwd = 2)

# Menambahkan legend di kiri atas
legend("topleft", legend = c("Harga Aset SP500", "Simulasi Normal"),
       col = c("blue", "orange"), lwd = 2, bty = "n")  # bty = "n" menghilangkan kotak di sekitar legend

# Tentukan tanggal untuk data testing (misalnya 12 Juli 2024 - 19 Juli 2024)
tanggal_testing <- asset_sp500$Date[train_size:n]  # Data testing dari tanggal setelah training set

# Ambil harga aset aktual untuk data testing
harga_testing <- asset_sp500$GSPC.Close[train_size:n]

# Plot harga aktual data testing
plot(tanggal_testing, harga_testing, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga S&P500", 
     main = "Harga Aset S&P500 dan Rata-rata Simulasi Normal", 
     xaxt = "n")  # Menghilangkan grid dan menyesuaikan sumbu x

# Tambahkan garis untuk hasil simulasi Normal pada data testing
lines(tanggal_testing, average_path, col = "orange", lwd = 2)

# Menambahkan legend di kiri atas
legend("topleft", legend = c("Harga Aset SP500", "Simulasi Normal"),
       col = c("blue", "orange"), lwd = 2, bty = "n")  # bty = "n" menghilangkan kotak di sekitar legend


ylim_range <- range(c(harga_testing, average_path)) 
# Plot harga aktual data testing dengan rentang sumbu Y yang lebih besar
plot(tanggal_testing, harga_testing, type = "l", col = "blue", lwd = 2, 
     xlab = "Tanggal", ylab = "Harga S&P500", 
     main = "Harga Aset S&P500 dan Rata-rata Simulasi Normal", 
     xaxt = "n", ylim = ylim_range)  # Menyesuaikan rentang sumbu Y
# Tambahkan garis untuk hasil simulasi Normal pada data testing
lines(tanggal_testing, average_path, col = "orange", lwd = 2)

# Menambahkan legend di kiri atas
legend("topleft", legend = c("Harga Aset SP500", "Simulasi Normal"),
       col = c("blue", "orange"), lwd = 2, bty = "n")  # bty = "n" menghilangkan kotak di sekitar legend

#### RMSE MAPE ####
# Menghitung harga opsi Call untuk setiap baris di data testing
r <- 0.0455  # Suku bunga tahunan di tgl maturity
call_prices <- numeric(nrow(test_call_sp500))  # Inisialisasi vektor untuk menyimpan harga opsi

for (i in 1:nrow(test_call_sp500)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_call_sp500$Strike[i]
  days_to_maturity <- test_call_sp500$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations+1]
  payoff_call <- pmax(0, ST - K)
  
  # Menghitung harga opsi dengan diskon faktor
  call_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_call) 
}

# Cetak hasil harga opsi untuk setiap baris di data testing
print(call_prices)

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual
results_sp500_call <- data.frame(
  Time = test_call_sp500$Time,
  Actual_Call_Price = test_call_sp500$Last,
  Simulated_Call_Price = call_prices
)

# Cetak hasil perbandingan
print(results_sp500_call)

# Hitung RMSE
rmse_sp500_call <- sqrt(mean((results_sp500_call$Actual_Call_Price - results_sp500_call$Simulated_Call_Price)^2))
cat("RMSE:", rmse_sp500_call, "\n")

# Hitung MAPE
mape_sp500_call <- mean(abs((results_sp500_call$Actual_Call_Price - results_sp500_call$Simulated_Call_Price) / results_sp500_call$Actual_Call_Price)) * 100
cat("MAPE:", mape_sp500_call, "%\n")

# Menghitung harga opsi Put untuk setiap baris di data testing
r <- 0.0455  # Suku bunga tahunan di tgl maturity
put_prices <- numeric(nrow(test_put_sp500))  # Inisialisasi vektor untuk menyimpan harga opsi Put

for (i in 1:nrow(test_put_sp500)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_put_sp500$Strike[i]
  days_to_maturity <- test_put_sp500$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations+1]
  payoff_put <- pmax(0, K - ST)
  
  # Menghitung harga opsi Put dengan diskon faktor
  put_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_put)
}

# Cetak hasil harga opsi Put untuk setiap baris di data testing
print(put_prices)

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual
results_put_sp500 <- data.frame(
  Time = test_put_sp500$Time,
  Actual_Put_Price = test_put_sp500$Last,
  Simulated_Put_Price = put_prices
)

# Cetak hasil perbandingan
print(results_put_sp500)

# Hitung RMSE
rmse_put_sp500 <- sqrt(mean((results_put_sp500$Actual_Put_Price - results_put_sp500$Simulated_Put_Price)^2))
cat("RMSE:", rmse_put_sp500, "\n")

# Hitung MAPE
mape_put_sp500 <- mean(abs((results_put_sp500$Actual_Put_Price - results_put_sp500$Simulated_Put_Price) / results_put_sp500$Actual_Put_Price)) * 100
cat("MAPE:", mape_put_sp500, "%\n")

# Tentukan path file untuk menyimpan hasil
library(writexl)
file_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results MC/5daystomaturity/Hasil MC5DayTM SP500.xlsx"

# Menyimpan kedua data frame dalam satu file Excel dengan dua sheet
write_xlsx(list(
  "Put Option Results" = results_put_sp500,
  "Call Option Results" = results_sp500_call
), path = file_path)



#### CI 95% ####
r <- 0.0455
# Misalkan jumlah simulasi adalah 10.000
num_simulations <- nrow(price_paths)

# Matriks untuk menyimpan harga opsi Call langsung dari payoff
call_prices_matrix <- matrix(0, nrow = num_simulations, ncol = nrow(test_call_sp500))

for (i in 1:nrow(test_call_sp500)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_call_sp500$Strike[i]
  days_to_maturity <- test_call_sp500$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations + 1]  # Harga akhir simulasi
  payoff_call <- pmax(0, ST - K)  # Payoff Call = max(0, ST - K)
  
  # Menghitung harga opsi Call untuk setiap simulasi (bukan rata-rata)
  call_prices_matrix[, i] <- exp(-r * days_to_maturity) * payoff_call
}

# Output: call_prices_matrix berisi harga opsi Call untuk setiap simulasi dan setiap baris data

# Path untuk menyimpan file Excel
output_call_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Kumpulan_Price_SP500_5DTM.xlsx"

# Simpan call_prices_matrix ke Excel
write_xlsx(list("Call Prices" = as.data.frame(call_prices_matrix)), path = output_call_prices_path)
cat("call_prices_matrix telah disimpan ke:", output_call_prices_path, "\n")

## PUT ##
# Menghitung harga opsi Put menggunakan simulasi Monte Carlo
r <- 0.0455
num_simulations <- nrow(price_paths)
put_prices_matrix <- matrix(0, nrow = num_simulations, ncol = nrow(test_put_sp500))

for (i in 1:nrow(test_put_sp500)) {
  K <- test_put_sp500$Strike[i]
  days_to_maturity <- test_put_sp500$Maturity[i]
  
  ST <- price_paths[, num_observations + 1]  # Harga akhir simulasi
  payoff_put <- pmax(0, K - ST)  # Payoff Put = max(0, K - ST)
  
  put_prices_matrix[, i] <- exp(-r * days_to_maturity) * payoff_put
}

output_put_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Kumpulan_PricePUT_SP500_5DTM.xlsx"
write_xlsx(list("Put Prices" = as.data.frame(put_prices_matrix)), path = output_put_prices_path)
cat("put_prices_matrix telah disimpan ke:", output_put_prices_path, "\n")

#### Menggunakan Percentile ####
# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo
price_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Kumpulan_Price_SP500_5DTM.xlsx")

# 2. Membaca dataset Testing Data Call
test_call_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing.xlsx', sheet = "Testing Data")

# 3. Menghitung quantile 5% dan 95% untuk setiap kolom di price_data
quantiles <- apply(price_data, 2, function(column) {
  quantile(column, probs = c(0.05, 0.95), na.rm = TRUE)
})

# Konversi hasil quantile menjadi data frame
quantile_df <- as.data.frame(t(quantiles))
colnames(quantile_df) <- c("Percentile_5%", "Percentile_95%")

# 4. Tambahkan kolom `Actual_Price` ke quantile_df dari test_call_sp500$Last
# Pastikan jumlah kolom di quantile_df sesuai dengan jumlah kolom di price_data
if (nrow(quantile_df) == nrow(test_call_sp500)) {
  quantile_df$Actual_Price <- test_call_sp500$Last
} else {
  stop("Jumlah kolom di quantile_df dan test_call_sp500 tidak sesuai. Harap cek data.")
}

# 5. Tambahkan kolom pengecekan apakah `Actual_Price` berada dalam rentang quantile
quantile_df$InRange <- quantile_df$Actual_Price >= quantile_df$'Percentile_5%' & 
  quantile_df$Actual_Price <= quantile_df$'Percentile_95%'

# 6. Simpan hasil ke file Excel baru
write.xlsx(quantile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Percentile_SPCall_5DTM.xlsx")


# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo untuk Put
priceput_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Kumpulan_PricePUT_SP500_5DTM.xlsx")

# 2. Membaca dataset Testing Data Put
test_put_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500 Training Testing PUT.xlsx', sheet = "Testing Data Put")

# 3. Menghitung quantile 5% dan 95% untuk setiap kolom di priceput_data
quantiles_put <- apply(priceput_data, 2, function(column) {
  quantile(column, probs = c(0.05, 0.95), na.rm = TRUE)
})

# Konversi hasil quantile menjadi data frame
quantileput_df <- as.data.frame(t(quantiles_put))
colnames(quantileput_df) <- c("Percentile_5%", "Percentile_95%")

# 4. Tambahkan kolom `Actual_Price` ke quantileput_df dari test_put_sp500$Last
# Pastikan jumlah kolom di quantileput_df sesuai dengan jumlah kolom di priceput_data
if (nrow(quantileput_df) == nrow(test_put_sp500)) {
  quantileput_df$Actual_Price <- test_put_sp500$Last
} else {
  stop("Jumlah kolom di quantileput_df dan priceput_data tidak sesuai. Harap cek data.")
}

# 5. Tambahkan kolom pengecekan apakah `Actual_Price` berada dalam rentang quantile
quantileput_df$InRange <- quantileput_df$Actual_Price >= quantileput_df$'Percentile_5%' & 
  quantileput_df$Actual_Price <= quantileput_df$'Percentile_95%'

# 6. Simpan hasil ke file Excel baru
write.xlsx(quantileput_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Percentile_SPPut_5DTM.xlsx")

##### Cek percentile untuk harga aset #####
library(readxl)
library(dplyr)
library(ggplot2)

# 1. Baca data simulasi
pricepath_data <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Pricepath_SP500_5DTM.xlsx')

# 2. Baca data actual
asset_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500_Data.xlsx', sheet="Asset Price All")

# 3. Bagi data menjadi 60% training dan 40% testing
n <- nrow(asset_sp500)
train_size <- floor(0.8 * n)
train_data <- asset_sp500[1:train_size, ]
test_data <- asset_sp500[(train_size + 1):n, ]

# 4. Gabungkan baris terakhir training dan seluruh testing sebagai actual data
actual_data <- rbind(train_data[train_size, ], test_data)

# 5. Hitung percentile 5% dan 95% untuk setiap kolom di data simulasi
percentile_df <- data.frame(
  Column = names(pricepath_data),
  Percentile_5 = apply(pricepath_data, 2, function(x) quantile(x, 0.05, na.rm = TRUE)),
  Percentile_95 = apply(pricepath_data, 2, function(x) quantile(x, 0.95, na.rm = TRUE))
)

# 6. Bandingkan nilai actual dengan percentile
actual_prices <- actual_data$GSPC.Close  # Sesuaikan kolom actual
percentile_df$Actual_Price <- actual_prices

percentile_df$In_Range <- mapply(
  function(actual, p5, p95) actual >= p5 & actual <= p95,
  actual_prices,
  percentile_df$Percentile_5,
  percentile_df$Percentile_95
)

# 7. Plot percentiles dan harga aktual
tanggal_aset <- actual_data$Date  # Sesuaikan dengan nama kolom tanggal
tanggal_simulasi <- tanggal_aset[(train_size):n]  # Rentang tanggal untuk data testing

ggplot(data = NULL, aes(x = tanggal_aset)) +
  geom_line(aes(y = actual_data$GSPC.Close, color = "Harga Aset SP500"), size = 1) +
  geom_line(aes(y = percentile_df$Percentile_5, color = "Percentile 5%"), linetype = "dashed") +
  geom_line(aes(y = percentile_df$Percentile_95, color = "Percentile 95%"), linetype = "dashed") +
  labs(
    title = "Harga Aset SP500 (Testing 20%) dan Percentile 5%-95%",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  theme_minimal()

ggplot() +
  # Plot harga aset untuk training data
  geom_line(data = train_data, aes(x = Date, y = GSPC.Close, color = "Training Data"), size = 1) +
  # Plot harga aset untuk testing data
  geom_line(data = test_data, aes(x = Date, y = GSPC.Close, color = "Testing Data"), size = 1) +
  
  # Label dan tampilan
  labs(
    title = "Harga Aset SP500 Rasio 80:20 (Training vs Testing)",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  scale_color_manual(values = c("Training Data" = "blue", "Testing Data" = "red")) +
  theme_minimal()

# 8. Simpan data hasil dengan percentile dan status
write.xlsx(percentile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP8020/Percentile_Result_AsetSP500_8020.xlsx")



##### Cek percentile untuk harga aset #####
library(readxl)
library(dplyr)
library(ggplot2)

# 1. Baca data simulasi
pricepath_data <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Pricepath_SP500_5DTM.xlsx')

# 2. Baca data actual
asset_sp500 <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/S&P500/SP500_Data.xlsx', sheet="Asset Price All")

# Ubah kolom tanggal menjadi format Date
asset_sp500$Date <- as.Date(asset_sp500$Date, format = "%Y-%m-%d")

# Pisahkan data training (sebelum 15 Juli 2024) dan data testing (15-19 Juli 2024)
train_data <- asset_sp500 %>% 
  filter(Date < as.Date("2024-07-15"))
test_data <- asset_sp500 %>% 
  filter(Date >= as.Date("2024-07-15") & Date <= as.Date("2024-07-19"))

# 4. Gabungkan baris terakhir training dan seluruh testing sebagai actual data
last_train_row <- train_data[nrow(train_data), ]  # Ambil baris terakhir dari training data
actual_data <- rbind(last_train_row, test_data)   # Gabungkan dengan seluruh data testing

# 5. Hitung percentile 5% dan 95% untuk setiap kolom di data simulasi
percentile_df <- data.frame(
  Column = names(pricepath_data),
  Percentile_5 = apply(pricepath_data, 2, function(x) quantile(x, 0.05, na.rm = TRUE)),
  Percentile_95 = apply(pricepath_data, 2, function(x) quantile(x, 0.95, na.rm = TRUE))
)

# 6. Bandingkan nilai actual dengan percentile
actual_prices <- actual_data$GSPC.Close # Sesuaikan kolom actual
percentile_df$Actual_Price <- actual_prices

percentile_df$In_Range <- mapply(
  function(actual, p5, p95) actual >= p5 & actual <= p95,
  actual_prices,
  percentile_df$Percentile_5,
  percentile_df$Percentile_95
)

# 7. Plot percentiles dan harga aktual
tanggal_aset <- actual_data$Date  # Sesuaikan dengan nama kolom tanggal
tanggal_simulasi <- actual_data$Date # Rentang tanggal untuk data testing

ggplot(data = NULL, aes(x = tanggal_aset)) +
  geom_line(aes(y = actual_data$GSPC.Close, color = "Harga Aset SP500"), size = 1) +
  geom_line(aes(y = percentile_df$Percentile_5, color = "Percentile 5%"), linetype = "dashed") +
  geom_line(aes(y = percentile_df$Percentile_95, color = "Percentile 95%"), linetype = "dashed") +
  labs(
    title = "Harga Aset SP500 (Testing 5 Days to Maturity) dan Percentile 5%-95%",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  theme_minimal()

ggplot() +
  # Plot harga aset untuk training data
  geom_line(data = train_data, aes(x = Date, y = GSPC.Close, color = "Training Data"), size = 1) +
  # Plot harga aset untuk testing data
  geom_line(data = test_data, aes(x = Date, y = GSPC.Close, color = "Testing Data"), size = 1) +
  
  # Label dan tampilan
  labs(
    title = "Harga Aset SP500 Rasio 5 Days to Maturity (Training vs Testing)",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  scale_color_manual(values = c("Training Data" = "blue", "Testing Data" = "red")) +
  theme_minimal()

# 8. Simpan data hasil dengan percentile dan status
write.xlsx(percentile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/SP5DTM/Percentile_Result_AsetSP500_5DTM.xlsx")
