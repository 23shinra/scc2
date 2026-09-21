import Foundation
import Vision
import AppKit

let path = CommandLine.arguments[1]
let url = URL(fileURLWithPath: path)
guard let img = NSImage(contentsOf: url),
      let tiff = img.tiffRepresentation,
      let rep = NSBitmapImageRep(data: tiff),
      let cg = rep.cgImage else {
    fputs("failed to load image\n", stderr)
    exit(1)
}

let req = VNRecognizeTextRequest()
req.recognitionLevel = .accurate
req.recognitionLanguages = ["ru-RU", "en-US"]
req.usesLanguageCorrection = true
if #available(macOS 13.0, *) {
    req.automaticallyDetectsLanguage = false
}
req.customWords = [
    "Северная", "Южная", "зона", "Западная", "Мангистауская",
    "Атырауская", "Казахстанская", "Всего", "МВт", "НДС",
    "Реестр", "сделок", "ТОО", "АО"
]

let handler = VNImageRequestHandler(cgImage: cg, options: [:])
try handler.perform([req])
let observations = req.results ?? []
for obs in observations {
    guard let c = obs.topCandidates(1).first else { continue }
    let b = obs.boundingBox
    let text = c.string.replacingOccurrences(of: "\t", with: " ").replacingOccurrences(of: "\n", with: " ")
    print(String(format: "%.5f\t%.5f\t%.5f\t%.5f\t%.3f\t%@", b.minX, b.maxX, b.minY, b.maxY, c.confidence, text))
}
