// swift-tools-version:5.9
import PackageDescription

let package = Package(
    name: "docx2pages",
    platforms: [
        .macOS(.v12)
    ],
    dependencies: [
        .package(url: "https://github.com/apple/swift-argument-parser", from: "1.2.0")
    ],
    targets: [
        .executableTarget(
            name: "docx2pages",
            dependencies: [
                .product(name: "ArgumentParser", package: "swift-argument-parser")
            ],
            path: "Sources/docx2pages",
            resources: [
                .copy("../../scripts")
            ]
        )
    ]
)
