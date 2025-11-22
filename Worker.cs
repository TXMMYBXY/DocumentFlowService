namespace DocumentFlowService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            var service = new DocumentTemplateService();

            var values = new Dictionary<string, string>
            {
                ["FullName"] = "Иванов Иван Иванович",
                ["LastnameFO"] = "Иванов И.И.",
                ["Datetime"] = DateTime.Now.ToString("dd.MM.yyyy")
            };

            var result = service.FillTemplate(@"F:\TestTemplate.docx", values);

            File.WriteAllBytes(@"F:\TestTemplate_Готовое.docx", result);

            while (!stoppingToken.IsCancellationRequested)
            {
                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                }
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}
