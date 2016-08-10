require "zip"
require "fileutils"
require "rexml/document"

md = ""
dest = "_ppt2md_tmp/"

Zip::File.open("#{ARGV[0]}") do |zip|
	zip.each do |entry|
		dirname = File::dirname entry.to_s
		if dirname == "ppt/slides"
			FileUtils.mkdir_p(dest + dirname)
			# puts "extract #{entry.to_s}"
			zip.extract(entry, dest + entry.to_s) {true}
		end
	end
end

i = 1
while File.exist?(dest + "ppt/slides/slide" + i.to_s + ".xml")
	xml = REXML::Document.new(open(dest + "ppt/slides/slide" + i.to_s + ".xml"))
	xml.elements.each('p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p') do |p|
		p.elements.each('a:r/a:t') do |t|
			print t.text
		end
		print "\n"
	end
	i = i + 1
end

#File.write("#{ARGV[1]}", md)

