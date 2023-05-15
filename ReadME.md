Automação para diariamente as 12 horas na hora do almoço, o script pegar o relaáorio de contador geral de impressões, e também de contador por usuários. O relátorio vem uma planilha com o formato .CSV, no script contem uma função que converte .csv para .xlsx.

Bibliotecas necessarias:
PyAutoGui
PyperClip
Pandas
datetime
time
schedule


Com o Scrpit que foi feito, o arquivo .bat, é só colocar no agendador do windowns para ele disparar o schedule no horário, no meu caso coloquei para o agendador da tarefas executar o scrpit as 11:50 AM, e de 12PM de acordo com o Schedule que foi configurado vai começar o primeiro processo.