Gem::Specification.new do |s|
  s.name          = 'jruby_excelcom'
  s.version       = '0.0.3'
  s.date          = Time.now.strftime('%Y-%m-%d')
  s.platform      = 'java'
  s.summary       = "Excel spreadsheet modification using COM"
  s.description   = "Uses the java library excelcom and JNA for modifying excel spreadsheets. Works on windows only."
  s.authors       = ["lprc"]
  s.files         = Dir.glob("{doc,lib,test}/**/*") + ['LICENSE', __FILE__]
  s.require_paths = ['lib']
  s.homepage      = 'https://github.com/lprc/jruby_excelcom'
  s.license       = 'Apache-2.0'
end