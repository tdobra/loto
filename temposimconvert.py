import subprocess

f = "TCMapCards.pdf"
num_stations = int(input("Number of stations: "))
num_problems = int(input("Tasks per station: "))
base = 1

fname = "tmp.pdf"
cname = "tmp-crop.pdf"

for s in range(1, num_stations+1):
    for p in range(1, num_problems+1):
        page = base + (s-1) * num_problems + (p-1)
        subprocess.run(["gs", "-dNOPAUSE", "-dBATCH", "-sOutputFile=%s" % fname, "-dFirstPage=%d" % page, "-dLastPage=%d" % page, "-sDEVICE=pdfwrite", "%s" % f])
        subprocess.run(["pdfcrop", "%s" % fname])
        outfile = "map-%d.%dz.png" % (s, p)
        subprocess.run(["convert", "-density", "150", "-quality", "100", "%s" % cname, "%s" % outfile])
        subprocess.run(["rm", "%s" % cname, "%s" % fname])
