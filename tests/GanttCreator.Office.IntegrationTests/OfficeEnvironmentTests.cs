using Xunit;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Tests that verify the Office integration environment is available.
/// These tests are tagged [Trait("Category","OfficeIntegration")] and
/// are excluded from verify-quick.ps1 / verify.ps1 on machines without
/// Microsoft 365 installed.
///
/// A failure here means the self-hosted runner setup is incomplete;
/// it does not mean the product is broken.
/// </summary>
public class OfficeEnvironmentTests
{
    private readonly ITestOutputHelper _output;

    public OfficeEnvironmentTests(ITestOutputHelper output) => _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public void Excel_interop_type_is_registered()
    {
        // Verify that the primary Excel interop assembly is registered
        // on this machine. This is a prerequisite for any real Office
        // automation tests.
        var excelType = Type.GetTypeFromProgID("Excel.Application");
        Assert.NotNull(
            excelType ?? throw new InvalidOperationException(
                "Excel.Application COM type is not registered. " +
                "Install Microsoft 365 or Microsoft Office for Windows " +
                "to run Office integration tests."));
        _output.WriteLine($"Excel COM type found: {excelType.FullName}");
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public void PowerPoint_interop_type_is_registered()
    {
        var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
        Assert.NotNull(
            pptType ?? throw new InvalidOperationException(
                "PowerPoint.Application COM type is not registered. " +
                "Install Microsoft 365 or Microsoft Office for Windows " +
                "to run Office integration tests."));
        _output.WriteLine($"PowerPoint COM type found: {pptType.FullName}");
    }
}
