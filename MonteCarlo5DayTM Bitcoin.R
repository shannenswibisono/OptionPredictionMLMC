# Load libraries
library(readxl)
library(dplyr)
library(ghyp)
library(stats)
library(ggplot2)
#library(fitdistrplus)
library(writexl)

# Read data untuk Bitcoin
asset_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Asset Price R")
test_call_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx', sheet = "Testing Data Call")
test_put_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx', sheet = "Testing Data Put")

# Ubah kolom tanggal menjadi format Date
asset_bitcoin$Time <- as.Date(asset_bitcoin$Time, format = "%Y-%m-%d")
test_call_btc$Time <- as.Date(test_call_btc$Time, format = "%Y-%m-%d")
test_put_btc$Time <- as.Date(test_put_btc$Time, format = "%Y-%m-%d")

########### BARU #########
# Menghitung log return dari harga penutupan Bitcoin
returns_bitcoin <- diff(log(asset_bitcoin$Last)) 
returns_bitcoin <- data.frame(Date = asset_bitcoin$Time[-1], Return = returns_bitcoin)

# Pisahkan data training (sebelum 24 Juni 2024) dan data testing (24-28 Juni 2024)
train_returns <- returns_bitcoin %>%
  filter(Date < as.Date("2024-06-24"))
test_returns <- returns_bitcoin %>%
  filter(Date >= as.Date("2024-06-24") & Date <= as.Date("2024-06-28"))

# Pastikan data memiliki variasi yang cukup
if (nrow(train_returns) == 0 || nrow(test_returns) == 0) {
  stop("Data training atau testing kosong!")
}
if (sd(train_returns$Return) == 0) {
  stop("Variabilitas data training terlalu kecil, fitting tidak dapat dilakukan.")
}

# Fit distribusi Normal pada data training
library(ghyp)
fit_gh_train <- fit.ghypuv(train_returns$Return)

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
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_8020.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Simulated Returns" = simulated_returns_df), path = output_path)

# Konfirmasi penyimpanan selesai
cat("Simulated returns telah disimpan ke:", output_path, "\n")

# Path ke file Excel yang berisi simulated_returns
input_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_8020.xlsx"

# Membaca data simulated_returns dari Excel
simulated_returns <- read_excel(input_path, sheet = "Simulated Returns")

# Konversi simulated_returns ke matrix untuk digunakan dalam perhitungan
simulated_returns <- as.matrix(simulated_returns)

# Inisialisasi harga awal (menggunakan harga terakhir data training Bitcoin)
S0_bitcoin <- 64260  # Menggunakan harga terakhir pada training data

# Jalur harga berdasarkan hasil simulasi
price_paths <- matrix(NA, nrow = num_simulations, ncol = num_observations+1)
price_paths[, 1] <- S0_bitcoin

# Menghitung price path
for (sim in 1:num_simulations) {
  for (day in 2:(num_observations+1)) {
    price_paths[sim, day] <- price_paths[sim, day - 1] * exp(simulated_returns[sim, day - 1])
  }
}

# Konversi pricepaths ke data frame
price_paths_df <- as.data.frame(price_paths)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Pricepath_BTC_8020.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Pricepath" = price_paths_df), path = output_path)

# Plot untuk beberapa jalur harga simulasi hingga 20 kolom pertama
plot(1:6, price_paths[1, 1:6], type = "l", col = "blue", ylim = range(price_paths[, 1:6]), 
     xlab = "Hari", ylab = "Harga Bitcoin", main = "Simulasi Jalur Harga Bitcoin (6 Hari Terakhir)")
# Tambahkan jalur harga lain ke plot
for (sim in 2:1000) {  # Misalnya, plotkan 200 jalur simulasi pertama
  lines(1:6, price_paths[sim, 1:6], col = sample(colors(), 1))
}


# Menghitung rata-rata price path
average_path <- apply(price_paths, 2, mean)

# Tentukan tanggal untuk plot sesuai dengan data testing
tanggal_aset <- asset_bitcoin$Time

# Tentukan tanggal batas training dan testing
train_date <- as.Date("2024-06-24")
train_size <- which(tanggal_aset == train_date) - 1  # Ukuran training set (jumlah data sebelum 15 Juli 2024)
n <- length(tanggal_aset)  # Jumlah total data

# Pisahkan tanggal simulasi (untuk testing)
tanggal_simulasi <- tanggal_aset[(train_size):n]

# Membuat data frame untuk plot
comparison_prices <- data.frame(
  Date = tanggal_aset,
  Actual_Price = asset_bitcoin$Last,
  Average_Simulated_Price = c(rep(NA, train_size-1), average_path)  # Isi NA untuk data sebelum simulasi dimulai
)

# Plot harga aktual dan rata-rata hasil simulasi price paths
# Plot harga aktual aset Bitcoin
plot(tanggal_aset, asset_bitcoin$Last, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga Bitcoin", 
     main = "Harga Aset Bitcoin dan Rata-rata Simulasi GH", 
     xaxt = "n")  # Menghilangkan grid dan menyesuaikan sumbu x

# Tambahkan garis untuk hasil simulasi Normal
lines(tanggal_simulasi, average_path, col = "orange", lwd = 2)

# Menambahkan legend di kiri atas
legend("topleft", legend = c("Harga Aset Bitcoin", "Simulasi GH"),
       col = c("blue", "orange"), lwd = 2, bty = "n")  # bty = "n" menghilangkan kotak di sekitar legend

# Tentukan tanggal untuk data testing
tanggal_testing <- asset_bitcoin$Time[train_size:n]  # Data testing dari tanggal setelah training set

# Ambil harga aset aktual untuk data testing
harga_testing <- asset_bitcoin$Last[train_size:n]

# Plot harga aktual data testing
plot(tanggal_testing, harga_testing, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga Bitcoin", 
     main = "Harga Aset Bitcoin - Data Testing", 
     xaxt = "n")  # Menghilangkan grid dan menyesuaikan sumbu x

# Tambahkan garis untuk hasil simulasi Normal pada data testing
lines(tanggal_testing, average_path, col = "orange", lwd = 2)

# Menambahkan legend di kiri atas
legend("topleft", legend = c("Harga Aset Bitcoin", "Simulasi Normal"),
       col = c("blue", "orange"), lwd = 2, bty = "n")  # bty = "n" menghilangkan kotak di sekitar legend

#### RMSE MAPE ####
# Menghitung harga opsi Call untuk setiap baris di data testing
r <- 0.0461  # Suku bunga tahunan di tgl maturity
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
results_bitcoin_call <- data.frame(
  Time = test_call_btc$Time,
  Actual_Call_Price = test_call_btc$Last,
  Simulated_Call_Price = call_prices
)

# Cetak hasil perbandingan
print(results_bitcoin_call)

# Hitung RMSE
rmse_bitcoin_call <- sqrt(mean((results_bitcoin_call$Actual_Call_Price - results_bitcoin_call$Simulated_Call_Price)^2))
cat("RMSE:", rmse_bitcoin_call, "\n")

# Hitung MAPE
mape_bitcoin_call <- mean(abs((results_bitcoin_call$Actual_Call_Price - results_bitcoin_call$Simulated_Call_Price) / results_bitcoin_call$Actual_Call_Price)) * 100
cat("MAPE:", mape_bitcoin_call, "%\n")

# Menghitung harga opsi Put untuk setiap baris di data testing
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
results_put_bitcoin <- data.frame(
  Time = test_put_btc$Time,
  Actual_Put_Price = test_put_btc$Last,
  Simulated_Put_Price = put_prices
)

# Cetak hasil perbandingan
print(results_put_bitcoin)

# Hitung RMSE
rmse_put_bitcoin <- sqrt(mean((results_put_bitcoin$Actual_Put_Price - results_put_bitcoin$Simulated_Put_Price)^2))
cat("RMSE:", rmse_put_bitcoin, "\n")

# Hitung MAPE
mape_put_bitcoin <- mean(abs((results_put_bitcoin$Actual_Put_Price - results_put_bitcoin$Simulated_Put_Price) / results_put_bitcoin$Actual_Put_Price)) * 100
cat("MAPE:", mape_put_bitcoin, "%\n")

# Tentukan path file untuk menyimpan hasil
file_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results MC/Bitcoin Results MC.xlsx"

# Menyimpan kedua data frame dalam satu file Excel dengan dua sheet
write_xlsx(list(
  "Put Option Results" = results_put_bitcoin,
  "Call Option Results" = results_bitcoin_call
), path = file_path)









###### TEST #######
# Load libraries
library(readxl)
library(dplyr)
library(ghyp)
library(ggplot2)
library(writexl)

# Read data
asset_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Asset Price R")
test_call_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx', sheet = "Testing Data Call")
test_put_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Put.xlsx', sheet = "Testing Data Put")

# Ubah kolom tanggal menjadi format Date
asset_bitcoin$Time <- as.Date(asset_bitcoin$Time, format = "%Y-%m-%d")
test_call_btc$Time <- as.Date(test_call_btc$Time, format = "%Y-%m-%d")

# Menghitung log return dari underlying asset Bitcoin
returns_btc <- diff(log(asset_bitcoin$Last)) 
returns_btc <- data.frame(Date = asset_bitcoin$Time[-1], Return = returns_btc)

# Membagi data training (sebelum 24 Juni 2024) dan data testing (24-28 Juni 2024)
train_returns <- returns_btc %>% 
  filter(Date < as.Date("2024-06-24"))
test_returns <- returns_btc %>% 
  filter(Date >= as.Date("2024-06-24") & Date <= as.Date("2024-06-28"))

# Pastikan data memiliki variasi yang cukup
if (nrow(train_returns) == 0 || nrow(test_returns) == 0) {
  stop("Data training atau testing kosong!")
}
if (sd(train_returns$Return) == 0) {
  stop("Variabilitas data training terlalu kecil, fitting tidak dapat dilakukan.")
}

# Fit Generalized Hyperbolic distribution pada data training
fit_ghyp_train <- fit.ghypuv(train_returns$Return)

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
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_5DTM.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Simulated Returns" = simulated_returns_df), path = output_path)

# Konfirmasi penyimpanan selesai
cat("Simulated returns telah disimpan ke:", output_path, "\n")

# Path ke file Excel yang berisi simulated_returns
input_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Simulated_Return_BTC_5DTM.xlsx"

# Membaca data simulated_returns dari Excel
simulated_returns <- read_excel(input_path, sheet = "Simulated Returns")

# Konversi simulated_returns ke matrix untuk digunakan dalam perhitungan
simulated_returns <- as.matrix(simulated_returns)

# Inisialisasi harga awal dan time steps
S0_call <- 64260  # Harga akhir training data (21 Juni 2024)
price_paths <- matrix(NA, nrow = num_simulations, ncol = num_observations + 1)
price_paths[, 1] <- S0_call

# Menghitung price path berdasarkan hasil simulasi
for (sim in 1:num_simulations) {
  for (day in 2:(num_observations + 1)) {
    price_paths[sim, day] <- price_paths[sim, day - 1] * exp(simulated_returns[sim, day - 1])
  }
}

# Konversi pricepaths ke data frame
price_paths_df <- as.data.frame(price_paths)

# Tentukan path untuk menyimpan file
output_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Pricepath_BTC_5DTM.xlsx"

# Simpan data frame ke file Excel
write_xlsx(list("Pricepath" = price_paths_df), path = output_path)

# Menghitung rata-rata price path
average_path <- apply(price_paths, 2, mean)

# Membuat data frame untuk plot
comparison_prices <- data.frame(
  Date = asset_bitcoin$Time[(nrow(train_returns)+1):nrow(asset_bitcoin)],
  Actual_Price = asset_bitcoin$Last[(nrow(train_returns)+1):nrow(asset_bitcoin)],  # Harga aktual di data testing
  Average_Simulated_Price = average_path
)

# Plot harga aktual dan rata-rata hasil simulasi price paths
# Menentukan rentang sumbu Y agar tidak terpotong
ylim_range <- range(c(comparison_prices$Actual_Price, comparison_prices$Average_Simulated_Price))

# Plot harga aset Bitcoin aktual dan rata-rata simulasi GH
plot(comparison_prices$Date, comparison_prices$Actual_Price, type = "l", col = "blue", lwd = 2,
     xlab = "Tanggal", ylab = "Harga Bitcoin", main = "Harga Aset Bitcoin dan Rata-rata Simulasi GH",
     ylim = ylim_range)  # Menentukan rentang sumbu Y

# Menambahkan garis simulasi GH
lines(comparison_prices$Date, comparison_prices$Average_Simulated_Price, col = "orange", lwd = 2)

# Menambahkan legenda dengan ukuran font lebih kecil
legend("topleft", legend = c("Harga Aset Bitcoin", "Simulasi GH"),
       col = c("blue", "orange"), lwd = 2, cex = 0.8)  # cex < 1 untuk mengecilkan ukuran legend


library(ggplot2)

# Membuat tanggal untuk plot
dates_training <- asset_bitcoin$Time[1:train_size]  # Tanggal periode training
dates_testing <- test_returns$Date  # Tanggal periode testing

# Gabungkan tanggal training dengan tanggal testing
dates_full <- c(dates_training, dates_testing)  # Semua tanggal dari periode training + testing

# Gabungkan harga aktual untuk seluruh data dan simulasi untuk periode testing
actual_prices <- asset_bitcoin$Last  # Semua harga aktual Bitcoin
simulated_prices_full <- c(rep(NA, train_size), average_path)  # Simulasi hanya diisi untuk periode testing

# Membuat data frame untuk plot
full_data <- data.frame(
  Date = dates_full,
  Actual_Price = actual_prices,
  Simulated_Price = simulated_prices_full
)

# Plot harga aktual dan simulasi GH
ggplot(full_data, aes(x = Date)) +
  # Plot harga aktual
  geom_line(aes(y = Actual_Price, color = "Harga Aset Bitcoin"), size = 1.2) +
  # Plot simulasi GH
  geom_line(aes(y = Simulated_Price, color = "Simulasi GH"), size = 1.2, linetype = "dashed") +
  # Penyesuaian plot
  scale_color_manual(values = c("blue", "orange")) +
  labs(title = "Harga Aset Bitcoin dan Rata-rata Simulasi GH",
       x = "Tanggal", y = "Harga Bitcoin") +
  theme_minimal() +
  theme(legend.position = "topleft", legend.title = element_blank())  # Menempatkan legend di kiri atas


#### RMSE MAPE ####
# Menghitung harga opsi Call untuk setiap baris di data testing
r <- 0.0461  # Suku bunga tahunan
call_prices <- numeric(nrow(test_call_btc))  # Inisialisasi vektor untuk menyimpan harga opsi

for (i in 1:nrow(test_call_btc)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_call_btc$Strike[i]
  days_to_maturity <- test_call_btc$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations + 1]
  payoff_call <- pmax(0, ST - K)
  
  # Menghitung harga opsi dengan diskon faktor
  call_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_call)
}

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual
results_btc_call <- data.frame(
  Time = test_call_btc$Time,
  Actual_Call_Price = test_call_btc$Last,
  Simulated_Call_Price = call_prices
)

# Hitung RMSE untuk opsi Call
rmse_btc_call <- sqrt(mean((results_btc_call$Actual_Call_Price - results_btc_call$Simulated_Call_Price)^2))
cat("RMSE Call:", rmse_btc_call, "\n")

# Hitung MAPE untuk opsi Call
mape_btc_call <- mean(abs((results_btc_call$Actual_Call_Price - results_btc_call$Simulated_Call_Price) / results_btc_call$Actual_Call_Price)) * 100
cat("MAPE Call:", mape_btc_call, "%\n")

# Menghitung harga opsi Put untuk setiap baris di data testing
put_prices <- numeric(nrow(test_put_btc))  # Inisialisasi vektor untuk menyimpan harga opsi Put

for (i in 1:nrow(test_put_btc)) {
  # Mengambil strike price dan time to maturity dari setiap baris di data testing
  K <- test_put_btc$Strike[i]
  days_to_maturity <- test_put_btc$Maturity[i]
  
  # Menghitung payoff berdasarkan harga akhir simulasi
  ST <- price_paths[, num_observations + 1]
  payoff_put <- pmax(0, K - ST)
  
  # Menghitung harga opsi Put dengan diskon faktor
  put_prices[i] <- exp(-r * days_to_maturity) * mean(payoff_put)
}

# Membuat DataFrame hasil perbandingan harga opsi simulasi dan harga aktual untuk opsi Put
results_put <- data.frame(
  Time = test_put_btc$Time,
  Actual_Put_Price = test_put_btc$Last,
  Simulated_Put_Price = put_prices
)

# Hitung RMSE untuk opsi Put
rmse_btc_put <- sqrt(mean((results_put$Actual_Put_Price - results_put$Simulated_Put_Price)^2))
cat("RMSE Put:", rmse_btc_put, "\n")

# Hitung MAPE untuk opsi Put
mape_btc_put <- mean(abs((results_put$Actual_Put_Price - results_put$Simulated_Put_Price) / results_put$Actual_Put_Price)) * 100
cat("MAPE Put:", mape_btc_put, "%\n")

# Menyimpan hasil simulasi dan harga opsi ke dalam file Excel
library(writexl)
file_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Results MC/5daystomaturity/Hasil MC5DayTM BTC.xlsx"
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
output_call_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_Price_BTC_5DTM.xlsx"

# Simpan call_prices_matrix ke Excel
write_xlsx(list("Call Prices" = as.data.frame(call_prices_matrix)), path = output_call_prices_path)
cat("call_prices_matrix telah disimpan ke:", output_call_prices_path, "\n")

# ---------------------------------------
# Proses membagi data per hari dan menghitung rata-rata
# ---------------------------------------

# Path file Excel yang berisi call_prices_matrix
input_call_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_Price_BTC_5DTM.xlsx"

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
output_average_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Average_Prices_BTC_5DTM.xlsx"

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
output_normality_results_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Normality_Results_BTC_5DTM.xlsx"

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
output_ci_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Confidence_Intervals_Per_Day_BTC_5DTM.xlsx"

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

output_put_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_PricePUT_BTC_5DTM.xlsx"
write_xlsx(list("Put Prices" = as.data.frame(put_prices_matrix)), path = output_put_prices_path)
cat("put_prices_matrix telah disimpan ke:", output_put_prices_path, "\n")

# Membaca data dan menghitung rata-rata harga per grup
input_put_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Kumpulan_PricePUT_BTC_5DTM.xlsx"
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

output_average_prices_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Average_PricesPUT_BTC_5DTM.xlsx"
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
output_normality_results_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Normality_Results_PUT_BTC_5DTM.xlsx"
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
output_ci_path <- "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/Confidence_Intervals_PUT_Per_Day_BTC_5DTM.xlsx"
write_xlsx(list("Confidence Intervals" = confidence_intervals_per_day), path = output_ci_path)
cat("Confidence intervals per day telah disimpan ke:", output_ci_path, "\n")


#### Menggunakan Percentile ####
# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo
price_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Kumpulan_Price_BTC_5DTM.xlsx")

# 2. Membaca dataset Testing Data Call
test_call_btc <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/Bitcoin Training Testing Call.xlsx', sheet = "Testing Data Call")

# 3. Menghitung quantile 5% dan 95% untuk setiap kolom di price_data
quantiles <- apply(price_data, 2, function(column) {
  quantile(column, probs = c(0.05, 0.95), na.rm = TRUE)
})

# Konversi hasil quantile menjadi data frame
quantile_df <- as.data.frame(t(quantiles))
colnames(quantile_df) <- c("Percentile_5%", "Percentile_95%")

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
write.xlsx(quantile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Percentile_BTCCall_5DTM.xlsx")


# Load library yang diperlukan
library(readxl)
library(openxlsx)

# 1. Membaca dataset Monte Carlo untuk Put
priceput_data <- read_excel("/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Kumpulan_PricePUT_BTC_5DTM.xlsx")

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
write.xlsx(quantileput_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Percentile_BTCPut_5DTM.xlsx")

##### Cek percentile untuk harga aset #####
library(readxl)
library(dplyr)
library(ggplot2)

# 1. Baca data simulasi
pricepath_data <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Pricepath_BTC_5DTM.xlsx')

# 2. Baca data actual
asset_bitcoin <- read_excel('/Users/shannenwibisono/Desktop/-SKRIPSI-/Bitcoin/BTM24 OPSI FIX.xlsx', sheet = "Asset Price R")

# Ubah kolom tanggal menjadi format Date
asset_bitcoin$Time <- as.Date(asset_bitcoin$Time, format = "%Y-%m-%d")

# Membagi data training (sebelum 24 Juni 2024) dan data testing (24-28 Juni 2024)
train_data <- asset_bitcoin %>% 
  filter(Time < as.Date("2024-06-24"))
test_data <- asset_bitcoin %>% 
  filter(Time >= as.Date("2024-06-24") & Time <= as.Date("2024-06-28"))

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
tanggal_simulasi <- actual_data$Time # Rentang tanggal untuk data testing

ggplot(data = NULL, aes(x = tanggal_aset)) +
  geom_line(aes(y = actual_data$Last, color = "Harga Aset BTC"), size = 1) +
  geom_line(aes(y = percentile_df$Percentile_5, color = "Percentile 5%"), linetype = "dashed") +
  geom_line(aes(y = percentile_df$Percentile_95, color = "Percentile 95%"), linetype = "dashed") +
  labs(
    title = "Harga Aset BTC (Testing 5 Days to Maturity) dan Percentile 5%-95%",
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
    title = "Harga Aset BTC Rasio 5 Days to Maturity (Training vs Testing)",
    x = "Tanggal",
    y = "Harga",
    color = "Legenda"
  ) +
  scale_color_manual(values = c("Training Data" = "blue", "Testing Data" = "red")) +
  theme_minimal()

# 8. Simpan data hasil dengan percentile dan status
write.xlsx(percentile_df, "/Users/shannenwibisono/Desktop/-SKRIPSI-/Monte Carlo/BTC5DTM/Percentile_Result_AsetBTC_5DTM.xlsx")
