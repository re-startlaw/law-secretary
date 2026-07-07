#!/usr/bin/env swift
import AppKit
import Foundation
import Vision

struct OCRResult: Encodable {
    let path: String
    let text: String
    let confidence: Float?
    let error: String
}

func emit(_ result: OCRResult) {
    let encoder = JSONEncoder()
    encoder.outputFormatting = [.withoutEscapingSlashes]
    if let data = try? encoder.encode(result), let line = String(data: data, encoding: .utf8) {
        print(line)
    }
}

func cgImage(from path: String) -> CGImage? {
    guard let image = NSImage(contentsOfFile: path) else {
        return nil
    }
    var rect = NSRect(origin: .zero, size: image.size)
    return image.cgImage(forProposedRect: &rect, context: nil, hints: nil)
}

func recognize(path: String) -> OCRResult {
    guard let image = cgImage(from: path) else {
        return OCRResult(path: path, text: "", confidence: nil, error: "image-open-failed")
    }

    var recognized: [String] = []
    var confidences: [Float] = []
    var requestError = ""

    let request = VNRecognizeTextRequest { request, error in
        if let error = error {
            requestError = "vision-error:\(type(of: error))"
            return
        }
        guard let observations = request.results as? [VNRecognizedTextObservation] else {
            requestError = "vision-no-results"
            return
        }
        for observation in observations {
            guard let candidate = observation.topCandidates(1).first else {
                continue
            }
            let text = candidate.string.trimmingCharacters(in: .whitespacesAndNewlines)
            if !text.isEmpty {
                recognized.append(text)
                confidences.append(candidate.confidence)
            }
        }
    }

    request.recognitionLevel = .accurate
    request.recognitionLanguages = ["ja-JP", "en-US"]
    request.usesLanguageCorrection = true

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do {
        try handler.perform([request])
    } catch {
        return OCRResult(path: path, text: "", confidence: nil, error: "vision-perform-error:\(type(of: error))")
    }

    let confidence = confidences.isEmpty ? nil : confidences.reduce(0, +) / Float(confidences.count)
    return OCRResult(path: path, text: recognized.joined(separator: "\n"), confidence: confidence, error: requestError)
}

let paths = CommandLine.arguments.dropFirst()
if paths.isEmpty {
    fputs("usage: evidence_vision_ocr.swift IMAGE...\n", stderr)
    exit(2)
}

for path in paths {
    emit(recognize(path: path))
}
