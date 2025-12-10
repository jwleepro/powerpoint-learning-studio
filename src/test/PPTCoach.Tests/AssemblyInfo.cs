using Xunit;

// Disable parallel test execution for PowerPoint COM automation tests
// Multiple tests accessing PowerPoint concurrently causes COM RPC failures
[assembly: CollectionBehavior(DisableTestParallelization = true)]
