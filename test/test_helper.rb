# encoding: utf-8

$:.unshift "#{File.expand_path(__FILE__)}/../lib"

# define public_send for Object for minitest-reporters (actually first defined in Ruby 1.9)
class Object
  alias :public_send :send unless method_defined?(:public_send)
end

require 'rubygems'
require 'minitest/reporters'
MiniTest::Reporters.use!
