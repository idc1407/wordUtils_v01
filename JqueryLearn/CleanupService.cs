public class CleanUpService : IHostedService, IDisposable
{
    private readonly ILogger _logger;
    private Timer _timer;

    public CleanUpService(ILogger<CleanUpService> logger)
    {
        _logger = logger;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Timed Background Service is starting.");

        _timer = new Timer(DoWork, null, TimeSpan.Zero,
            TimeSpan.FromMinutes(1));

        return Task.CompletedTask;
    }

    private void DoWork(object state)
    {

        string targetDirectory = @"d:\itemp\_delproc";
        string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);

        foreach (string subdirectory in subdirectoryEntries)
        {
            DateTime currentTime = DateTime.Now;
            DateTime directoryTime = Directory.GetCreationTime(subdirectory);
            TimeSpan duration = currentTime - directoryTime;
            if(duration.Minutes > 5)
            {
                try
                {
                    Directory.Delete(subdirectory, true);
                    _logger.LogInformation(subdirectory + "  Deleted");
                }
                catch { }
            }
            
        }
        _logger.LogInformation("Timed Background Service is working.");
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Timed Background Service is stopping.");

        _timer?.Change(Timeout.Infinite, 0);

        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _timer?.Dispose();
    }
}