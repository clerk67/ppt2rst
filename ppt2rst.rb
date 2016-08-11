#! /usr/bin/env ruby -Ku

require "fileutils"
require "rexml/document"
require "zip"

SLIDES_DIR = "ppt/slides"
TEMP_DIR = ".ppt2rst_tmp/"
TITLELINE_LENGTH = 80
PROGRESSBAR_LENGTH = 63

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
		if dirname == SLIDES_DIR
			count += 1
		end
	end
	i = 1
	zip.each do |entry|
		dirname = File::dirname(entry.to_s)
		if dirname == SLIDES_DIR
			printf("\r[%3d/%3d] |", i, count)
			print "=" * (PROGRESSBAR_LENGTH * i / count) + " " * (PROGRESSBAR_LENGTH - (PROGRESSBAR_LENGTH * i / count))
			printf("| %3d\%", 100 * i / count)
			zip.extract(entry, TEMP_DIR + entry.to_s) {true}
			i += 1
		end
	end
end

count = 0
ls = Dir::entries(TEMP_DIR + SLIDES_DIR)
ls.each do |line|
	if line.match(/slide\d+\.xml/)
		count += 1
	end
end

i = 1
rst = ""
puts "\nGenerating reStructuredText..."

while File.exist?(TEMP_DIR + SLIDES_DIR + "/slide" + i.to_s + ".xml")
	printf("\r[%3d/%3d] |", i, count)
	print "=" * (PROGRESSBAR_LENGTH * i / count) + " " * (PROGRESSBAR_LENGTH - (PROGRESSBAR_LENGTH * i / count))
	printf("| %3d\%", 100 * i / count)
	xml = REXML::Document.new(open(TEMP_DIR + SLIDES_DIR + "/slide" + i.to_s + ".xml"))
	rst << ".. slide #{i}/#{count}\n\n"
	spTree = xml.elements["p:sld/p:cSld/p:spTree"]
	spTree.elements.each do |element|
		case element.name
		when "sp"
			sp = element
			ph = sp.elements["p:nvSpPr/p:nvPr/p:ph"]
			heading = 0
			unless ph.nil?
				case ph.attributes["type"]
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
				pPr = p.elements["a:pPr"]
				bullet = "* "
				level = 0
				unless pPr.nil?
					unless pPr.elements["a:buNone"].nil?
						bullet = ""
					else
						unless pPr.elements["a:buAutoNum"].nil?
							bullet = "#. "
						end
						lvl = pPr.attributes["lvl"]
						unless lvl.nil?
							level = lvl.to_i
						end
					end
				end
				if heading == 0
					rst << "\t" * level + bullet
				end
				p.elements.each("a:r") do |r|
					rPr = r.elements["a:rPr"]
					astah = 0
					unless rPr.nil?
						if rPr.attributes["i"] == "1"
							astah += 1
						end
						if rPr.attributes["b"] == "1"
							astah += 2
						end
					end
					t = r.elements["a:t"]
					unless t.nil? || t.text.to_s.empty?
						text = t.text.to_s.gsub(/\*/, "\\*").gsub(/\#/, "\\#")
						unless astah == 0 || rst[-1] == " "
							rst << " "
						end
						rst << "*" * astah
						if astah == 0
							rst << text
						else
							rst << text.strip
						end
						rst << "*" * astah
						unless astah == 0
							rst << " "
						end
					end
				end
				if heading == 0
					rst << "\n"
				end
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
			unless tbl.nil?
				rst << ".. list-table::\n\n"
				tbl.elements.each("a:tr") do |tr|
					rst << "\t*"
					tr.elements.each("a:tc") do |tc|
						if rst[-1] == "*"
							rst << " - "
						else
							rst << "\t  - "
						end
						tc.elements.each("a:txBody/a:p/a:r/a:t") do |t|
							rst << t.text.to_s.gsub(/\*/, "\\*").gsub(/\#/, "\\#")
						end
						rst << "\n"
					end
				end
				rst << "\n"
			end
		end
	end
	i += 1
end

File.write("#{ARGV[1]}", rst)
puts "\nSUCCESS: #{rst.count("\n")} lines were written to #{ARGV[1]}"

