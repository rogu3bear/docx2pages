# Homebrew formula for docx2pages
# To install from this repository as a tap:
#   brew tap yourusername/docx2pages https://github.com/yourusername/docx2pages
#   brew install docx2pages
#
# Or install directly from the formula:
#   brew install --formula Formula/docx2pages.rb

class Docx2pages < Formula
  desc "Convert Microsoft Word documents to Apple Pages format"
  homepage "https://github.com/rogu3bear/docx2pages"
  url "https://github.com/rogu3bear/docx2pages/archive/refs/tags/v1.4.0.tar.gz"
  # TODO: Update SHA256 after pushing v1.4.0 tag and creating release
  # Compute with: curl -sL https://github.com/rogu3bear/docx2pages/archive/refs/tags/v1.4.0.tar.gz | shasum -a 256
  sha256 "PLACEHOLDER_SHA256_NEEDS_UPDATE"
  license "MIT"
  head "https://github.com/rogu3bear/docx2pages.git", branch: "main"

  depends_on :macos
  depends_on "python@3"
  depends_on xcode: ["14.0", :build]

  def install
    # Build the Swift binary
    system "swift", "build", "-c", "release", "--disable-sandbox"

    # Install the binary
    bin.install ".build/release/docx2pages" => "docx2pages-bin"

    # Install the required scripts
    libexec.install "scripts/parse_docx.py"
    libexec.install "scripts/pages_writer.js"

    # Create a wrapper script that sets the scripts directory
    (bin/"docx2pages").write <<~EOS
      #!/bin/bash
      exec "#{bin}/docx2pages-bin" --scripts-dir "#{libexec}" "$@"
    EOS
  end

  def caveats
    <<~EOS
      docx2pages requires:
        - Apple Pages.app to be installed
        - Python package 'defusedxml': pip3 install defusedxml
        - macOS Automation permissions for Pages (granted on first run)

      First-time setup:
        1. Install defusedxml: pip3 install defusedxml
        2. Create a template: Open Pages, create a document with your preferred
           styles, and save it (e.g., ~/Documents/MyTemplate.pages)
        3. Run: docx2pages -i input.docx -o output.pages -t ~/Documents/MyTemplate.pages

      For more information, see: #{homepage}
    EOS
  end

  test do
    # Basic version check
    assert_match "docx2pages", shell_output("#{bin}/docx2pages --version")

    # Verify help output contains expected options
    help_output = shell_output("#{bin}/docx2pages --help")
    assert_match "--input", help_output
    assert_match "--output", help_output
    assert_match "--template", help_output
  end
end
