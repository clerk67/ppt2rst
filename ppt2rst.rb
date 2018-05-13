#!/usr/bin/env ruby -Ku

require "fileutils"
require "rexml/document"
require "zip"

SLIDES_DIR = "ppt/slides"
TEMP_DIR = ".ppt2rst_tmp/"
TITLELINE_LENGTH = 80
PROGRESSBAR_LENGTH = 63

def p2rst(p, heading)
  return "" if p.elements["a:r"].nil?
  pPr = p.elements["a:pPr"]
  bullet = "* "
  level = 0
  unless pPr.nil?
    bullet = "" unless pPr.elements["a:buNone"].nil?
    bullet = "#. " unless pPr.elements["a:buAutoNum"].nil?
    lvl = pPr.attributes["lvl"]
    level = lvl.to_i unless lvl.nil?
  end
  rst = ""
  rst << "\t" * level + bullet if heading == 0
  p.elements.each("a:r") do |r|
    rst << r2rst(r)
  end
  rst << "\n" if heading == 0
  rst
end

def r2rst(r)
  rPr = r.elements["a:rPr"]
  if rPr&.attributes["b"] == "1"
    astah = 2
  elsif rPr&.attributes["i"] == "1"
    astah = 1
  else
    astah = 0
  end
  rst = ""
  t = r.elements["a:t"]
  unless t&.text.to_s.empty?
    text = t.text.to_s.gsub(/\*/, "\\*").gsub(/\#/, "\\#")
    rst << "\\ " unless astah == 0 || rst.size == 0 || rst[-1] == " "
    rst << "*" * astah
    if astah == 0
      rst << text
    else
      rst << text.strip
    end
    rst << "*" * astah
    rst << " " unless astah == 0
  end
  rst
end

def tbl2rst(tbl, header_rows, stub_columns)
  rst = ".. list-table:: \\ \n"
  rst << "   :header-rows: #{header_rows}\n"
  rst << "   :stub-columns: #{stub_columns}\n\n"
  tbl.elements.each("a:tr") do |tr|
    rst << "   *"
    tr.elements.each("a:tc") do |tc|
      rst << "    " unless rst[-1] == "*"
      rst << " - "
      tc.elements.each("a:txBody/a:p/a:r") do |r|
        rst << r2rst(r)
      end
      rst << "\n"
    end
  end
  rst << "\n"
  rst
end

def show_progress(value, max)
  printf("\r[%3d/%3d] |", value, max)
  print "=" * (PROGRESSBAR_LENGTH * value / max) + " " * (PROGRESSBAR_LENGTH - (PROGRESSBAR_LENGTH * value / max))
  printf("| %3d\%", 100 * value / max)
end

unless ARGV.size == 2
  puts "ERROR: Invalid arguments"
  exit
end

unless File.exist?("#{ARGV[0]}")
  puts "ERROR: Input file not found"
  exit
end

puts "Extracting input files..."

FileUtils::mkdir_p(TEMP_DIR + SLIDES_DIR)
Zip::File.open("#{ARGV[0]}") do |zip|
  count = 0
  zip.each do |entry|
    dirname = File::dirname(entry.to_s)
    count += 1 if dirname == SLIDES_DIR
  end
  i = 0
  zip.each do |entry|
    dirname = File::dirname(entry.to_s)
    if dirname == SLIDES_DIR
      show_progress(i + 1, count)
      zip.extract(entry, TEMP_DIR + entry.to_s) {true}
      i += 1
    end
  end
end

count = 0
ls = Dir::entries(TEMP_DIR + SLIDES_DIR)
ls.each do |line|
  count += 1 if line.match(/slide\d+\.xml/)
end

rst = ""
puts "\nGenerating reStructuredText..."

count.times do |i|
  show_progress(i + 1, count)
  xml = REXML::Document.new(open(TEMP_DIR + SLIDES_DIR + "/slide" + (i + 1).to_s + ".xml"))
  rst << ".. slide #{i + 1}/#{count}\n\n"
  spTree = xml.elements["p:sld/p:cSld/p:spTree"]
  spTree.elements.each do |element|
    case element.name
    when "sp"
      sp = element
      nvSpPr = sp.elements["p:nvSpPr"]
      cNvSpPr = nvSpPr.elements["p:cNvSpPr"]
      heading = 0
      txBox = false
      if cNvSpPr&.attributes["txBox"] == "1"
        heading = -1
        txBox = true
      end
      ph = nvSpPr.elements["p:nvPr/p:ph"]
      unless ph.nil?
        case ph.attributes["type"]
        when "ctrTitle"
          heading = 1
        when "title"
          heading = 3
        when "subTitle"
          heading = 4
        when "hdr"
          heading = -1
        when "ftr"
          heading = -1
        when "sldNum"
          heading = -1
        end
      end
      case heading
      when -1
        next
      when 1
        rst << "=" * TITLELINE_LENGTH + "\n"
      when 3
        rst << "-" * TITLELINE_LENGTH + "\n"
      end
      sp.elements.each("p:txBody/a:p") do |p|
        rst << p2rst(p, heading)
      end
      rst << "\n"
      case heading
      when 1, 2
        rst << "=" * TITLELINE_LENGTH + "\n\n"
      when 3, 4
        rst << "-" * TITLELINE_LENGTH + "\n\n"
      end
    when "graphicFrame"
      graphicFrame = element
      tbl = graphicFrame.elements["a:graphic/a:graphicData/a:tbl"]
      rst << tbl2rst(tbl, 0, 0) unless tbl.nil?
    end
  end
end

File.write("#{ARGV[1]}", rst)
FileUtils.rm_rf(TEMP_DIR)
puts "\nSUCCESS: #{rst.count("\n")} lines were written to #{ARGV[1]}"

