require "zip"
require "fileutils"
require "rexml/document"

TITLELINE_LENGTH = 80

unless ARGV.size == 2
	puts "ERROR: Invalid arguments"
	exit
end

unless File.exist?("#{ARGV[0]}")
	puts "ERROR: Input file not found"
	exit
end

dest = "_ppt2md_tmp/"
Zip::File.open("#{ARGV[0]}") do |zip|
	zip.each do |entry|
		dirname = File::dirname entry.to_s
		if dirname == "ppt/slides"
			FileUtils.mkdir_p(dest + dirname)
			zip.extract(entry, dest + entry.to_s) {true}
		end
	end
end

i = 1
rst = ""
while File.exist?(dest + "ppt/slides/slide" + i.to_s + ".xml")
	xml = REXML::Document.new(open(dest + "ppt/slides/slide" + i.to_s + ".xml"))
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
				when "subtitle"
					heading = 4
				end
			end
			case heading
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
							bullet = "# "
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
					unless t.nil?
						rst << "*" * astah
						rst << t.text.to_s.gsub(/\*/, "\\*").gsub(/\#/, "\\#")
						rst << "*" * astah
					end
				end
				if heading == 0
					rst << "\n"
				end
			end
			rst << "\n"
			case heading
			when 1, 2
				rst << "=" * TITLELINE_LENGTH + "\n"
			when 3, 4
				rst << "-" * TITLELINE_LENGTH + "\n"
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
	i = i + 1
end

File.write("#{ARGV[1]}", rst.gsub(/\*{4}/, ""))

