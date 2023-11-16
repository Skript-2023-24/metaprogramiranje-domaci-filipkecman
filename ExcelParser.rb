require 'roo'
require 'spreadsheet'

class Kolona
  attr_accessor :naziv, :polja, :tabela

  def initialize(naziv, polja, tabela)
    @naziv = naziv
    @polja = polja
    @tabela = tabela
  end

  def inspect
    "#{naziv}: #{polja.inspect}"
  end

  def to_s
    "#{naziv}: #{polja.inspect}"
  end

  def map(&block)
    polja.map(&block)
  end

  def select(&block)
    polja.select(&block)
  end

  def reduce(initial = 0, &block)
    polja.reduce(initial, &block)
  end

  def suma
    polja.sum
  end

  def avg
    polja.sum.to_f/polja.length()
  end

  def []=(index, value)
    polja[index] = value
    update_tabela(index, value)
  end


  def update_tabela(index, value)
    @tabela.redovi.each do |red|
      red[index + 1] = value if red[0] == @naziv
    end
  end

  def [](n)
    polja[n]
  end

  def generisi_metode_celija
    @polja.each do |celija|
      define_singleton_method("#{celija}") do
        indeks = -1 
        @tabela.redovi.each_with_index do |red, i|
          if i == 0
            red.each_with_index do |celija_reda, j|
              indeks = j if celija_reda == @naziv
            end
          else
            red.each_with_index do |celija_reda, j|
              return red if celija_reda == celija && j == indeks
            end
          end
        end
      end
    end
  end
end
  
class Tabela
    attr_accessor :naziv, :kolone, :redovi
    include Enumerable

    def initialize(naziv)
      @naziv = naziv
      @kolone = []
      @redovi = []
    end

    def red(indeks)
      redovi[indeks]
    end

    def [](naziv_kolone)
      kolone.find {|col| col.naziv == naziv_kolone} || "Trazena kolona ne postoji!"
    end

    def to_s
      "#{redovi.map { |row| row.join(' ') }.join("\n")}"
    end


    def inicijalizuj_kolone
      kolone_tabele = @redovi.transpose
      kolone_tabele.each do |kol|
        trenutna_kolona = Kolona.new(kol[0], kol[1..-1], self)
        @kolone << trenutna_kolona
      end

      @kolone.each do |kol|
        kol_naziv = kol.naziv

        define_singleton_method("#{kol_naziv}") do
          @kolone.find { |kolona| kolona.naziv == kol_naziv}
        end
        kol.generisi_metode_celija
      end
    end
    

    def each(&block)
      @redovi.each do |red|
        @kolone.each do |kol|
          block.call(red[kolone.find_index(kol)])
        end
      end
    end
end


def inicijalizacija_tabela(naziv_fajla)
  tabele = []
  if naziv_fajla.end_with?(".xlsx")
    workbook = Roo::Spreadsheet.open(naziv_fajla, expand_merged_ranges: true)
    worksheets = workbook.sheets

    worksheets.each do |worksheet|
      broj_redova = 0
      tabele.append(Tabela.new("Tabela_#{worksheet}"))
      workbook.sheet(worksheet).each_row_streaming do |red|
        celije_reda = red.map(&:value)
        celije_reda_f = red.map { |celija| celija}
        tabele[-1].redovi.append(celije_reda) unless celije_reda_f.to_s.downcase.include?("subtotal") || celije_reda_f.to_s.downcase.include?("total")
        broj_redova += 1
      end
      tabele[-1].inicijalizuj_kolone
    end
  else
    workbook = Spreadsheet.open(naziv_fajla, expand_merged_ranges: true)
    worksheets = workbook.worksheets

    worksheets.each do |worksheet|
      broj_redova = 0
      tabele.append(Tabela.new("Tabela_#{worksheet}"))
      worksheet.redovi.each do |red|
        celije_reda = red.to_a.map { |v| v.methods.include?(:value) ? v : v}
        unless celije_reda.length.zero? || celije_reda.to_s.include?("Formula")
          #celije_reda.compact!
          tabele[-1].redovi.append(celije_reda)
        end
        broj_redova += 1
      end
      tabele[-1].inicijalizuj_kolone
    end
  end
  tabele
end






def main
  ime_fajla = '/media/sf_Shared_Kali/skript_jezici/cars.xlsx'
  tabele = inicijalizacija_tabela(ime_fajla)
  tabele.each do |tabela|
    puts tabela
    puts tabela.red(1)
    puts tabela["Marka"]
    puts tabela["Marka"][0]
    tabela["Marka"][0] = "Opel"
    puts tabela["Marka"][0]
    puts tabela.Marka
    puts tabela.Konjaza.suma
    puts tabela.Konjaza.avg
    puts tabela.Marka.BMW
    puts tabela.Konjaza.map { |x| x*2}
    puts tabela.Konjaza.select { |x| x > 200}
    puts tabela.Konjaza.select { |x| x > 215}
    puts tabela.Konjaza.select { |x| x > 250}
    p tabela.Konjaza.reduce { |sum,x| sum+x}
    puts tabela.Konjaza.avg
  end
end

main

  
