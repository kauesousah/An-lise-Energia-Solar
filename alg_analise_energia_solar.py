import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Criar diretório para salvar gráficos
OUTPUT_PATH = "graficos_gerados"
os.makedirs(OUTPUT_PATH, exist_ok=True)

FILE_NAME = 'Dados tratados.xlsx'

# Carrega o arquivo e processa os dados
def load_and_process_data(file_path):
  print("> Lendo arquivo de dados...")
  df = pd.read_excel(file_path, engine="openpyxl")
  df.rename(columns={'Ano': 'year', 'Mês': 'month', 'Dia': 'day', 'Hora': 'hour', 'Minuto': 'minute'}, inplace=True)
  
  print("> Processando dados...")
  # Criar coluna de data e hora
  df['datetime'] = pd.to_datetime(df[['year', 'month', 'day', 'hour', 'minute']].astype(str).agg('-'.join, axis=1), errors='coerce', format='%Y-%m-%d-%H-%M')
  
  # Calcular energia e eficiência
  df['energy_ac'] = (df['Vac[V]'] * df['Ica[A]'] * df['FP'] / 1000) * df['Período[h]']
  df['energy_dc'] = (df['Vcc1[V MPPT1]'] * df['Icc1[A]'] / 1000) * df['Período[h]']
  df['efficiency'] = np.where(df['energy_dc'] != 0, (df['energy_ac'] / df['energy_dc']) * 100, 0)
  df['power_ac'] = df['Vac[V]'] * df['Ica[A]'] * df['FP'] / 1000

  # Extrair data para análise diária
  df['date'] = df['datetime'].dt.date
  return df

# Calcula a energia diária
def calculate_daily_energy(df):
  return df.groupby('date')['energy_ac'].sum()

# Identifica os dias de produção máxima, média e mínima
def identify_production_days(daily_energy):
  max_day = daily_energy.idxmax()
  min_day = daily_energy.idxmin()
  mean_day = daily_energy.sub(daily_energy.mean()).abs().idxmin()
  return max_day, min_day, mean_day

# Mostra as informações de produção de energia
def display_energy_info(daily_energy, max_day, min_day, mean_day):
  print(f"Dia de máxima produção (CA): {max_day.strftime('%d/%m/%Y')}")
  print(f"Energia produzida: {daily_energy[max_day]:.2f} kWh\n")

  print(f"Dia de média produção (CA): {mean_day.strftime('%d/%m/%Y')}")
  print(f"Energia produzida: {daily_energy[mean_day]:.2f} kWh\n")

  print(f"Dia de mínima produção (CA): {min_day.strftime('%d/%m/%Y')}")
  print(f"Energia produzida: {daily_energy[min_day]:.2f} kWh\n")

# Plot gráficos de produção e eficiência
def plot_production(data, title, filename):
  fig, ax1 = plt.subplots(figsize=(10, 6))

  ax1.plot(data['datetime'], data['power_ac'], 'b-', label='Potência CA (kW)')
  ax1.set_xlabel('Horário')
  ax1.set_ylabel('Potência (kW)', color='b')
  ax1.tick_params('y', colors='b')

  ax2 = ax1.twinx()
  ax2.plot(data['datetime'], data['efficiency'], 'r-', label='Eficiência (%)')
  ax2.set_ylabel('Eficiência (%)', color='r')
  ax2.tick_params('y', colors='r')

  plt.title(title)
  fig.tight_layout()
  plt.savefig(filename)
  plt.show()
  plt.close()

# Plot e salva o gráfico de energia mensal
def plot_monthly_energy(monthly_energy, OUTPUT_PATH):
    fig, ax = plt.subplots(figsize=(10, 6))
    monthly_energy.unstack().plot(kind='bar', ax=ax)
    ax.set_title('Energia Mensal Produzida (CA)')
    ax.set_xlabel('Mês')
    ax.set_ylabel('Energia (kWh)')
    ax.legend(title='Ano')
    fig.tight_layout()
    fig.savefig(os.path.join(OUTPUT_PATH, 'energia_mensal.png'))
    plt.show()
    plt.close(fig)

# Plot e salva o gráfico de eficiência mensal
def plot_monthly_efficiency(monthly_efficiency, OUTPUT_PATH):
    fig, ax = plt.subplots(figsize=(10, 6))
    monthly_efficiency.unstack().plot(kind='bar', ax=ax)
    ax.set_title('Eficiência Mensal do Inversor')
    ax.set_xlabel('Mês')
    ax.set_ylabel('Eficiência (%)')
    ax.legend(title='Ano')
    fig.tight_layout()
    fig.savefig(os.path.join(OUTPUT_PATH, 'eficiencia_mensal.png'))
    plt.show()
    plt.close(fig)

def process_monthly_data(df):
  monthly_energy = df.groupby(['year', 'month'])['energy_ac'].sum()
  monthly_efficiency = df.groupby(['year', 'month'])['efficiency'].mean()
  return monthly_energy, monthly_efficiency
#
#
#
# Inicio
#

df = load_and_process_data(FILE_NAME)

daily_energy = calculate_daily_energy(df)

max_day, min_day, mean_day = identify_production_days(daily_energy)

# Visualizar dados
display_energy_info(daily_energy, max_day, min_day, mean_day)

# Processa os dados
monthly_energy, monthly_efficiency = process_monthly_data(df)

# CRIAÇÃO DE GRÁFICOS
# dias de produção máxima, média e mínima
plot_production(df[df['date'] == max_day], f'Produção Máxima ({max_day.strftime("%d/%m/%Y")})', os.path.join(OUTPUT_PATH, 'producao_maxima.png'))
plot_production(df[df['date'] == mean_day], f'Produção Média ({mean_day.strftime("%d/%m/%Y")})', os.path.join(OUTPUT_PATH, 'producao_media.png'))
plot_production(df[df['date'] == min_day], f'Produção Mínima ({min_day.strftime("%d/%m/%Y")})', os.path.join(OUTPUT_PATH, 'producao_minima.png'))

# Energia e eficiência mensal
plot_monthly_energy(monthly_energy, OUTPUT_PATH)
plot_monthly_efficiency(monthly_efficiency, OUTPUT_PATH)


