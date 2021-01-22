class CollectData
  def self.call(args)
    service = new(args)
    service.call
    service
  end

  def initialize(args)
    @args = args
  end

  private

  attr_reader :args
end
