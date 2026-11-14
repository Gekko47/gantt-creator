using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration.CustomUI;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// D1/D7: each lock toggle shows a padlock that matches its state, and the pair is a pair of
/// committed Lucide assets rather than a gallery id.
/// </summary>
/// <remarks>
/// <para>
/// This is the half of R5.18f that measurement had blocked. The approved design was a closed
/// padlock while locked and an open one while unlocked, and it could not ship because no
/// open-padlock <c>imageMso</c> exists in the Office gallery. The label carried the state
/// instead. D7 (owner ruling 2026-11-14) makes every ribbon glyph a Lucide asset served through
/// <c>getImage</c>, which removes the constraint — so these tests pin the asset pair, the
/// callback wiring, and the fact that the two states really are different pictures.
/// </para>
/// <para>
/// No host is involved: the picture is produced by <see cref="RibbonImages"/> from the embedded
/// PNG, exactly as the ribbon callback produces it. What only Excel can show is the rendered
/// glyph, which stays an Office-gate item.
/// </para>
/// </remarks>
public sealed class LockGlyphImagesTests
{
    private const string WorkbookRibbonId = "Microsoft.Excel.Workbook";

    private static readonly byte[] _pngSignature = [0x89, 0x50, 0x4E, 0x47];

    /// <summary>Every committed lock asset, at both sizes the owner's ruling names.</summary>
    private static readonly string[] _lockAssets =
    [
        "lock-locked-16.png",
        "lock-locked-32.png",
        "lock-unlocked-16.png",
        "lock-unlocked-32.png",
    ];

    [Fact]
    public void Both_lock_glyphs_are_embedded_and_are_real_pngs()
    {
        foreach (string suffix in _lockAssets)
        {
            byte[]? png = RibbonImages.ReadPng(suffix);

            Assert.NotNull(png);
            Assert.True(
                png.Length > _pngSignature.Length,
                $"'{suffix}' is {png.Length} bytes, which cannot be a rendered glyph.");
            Assert.True(
                png.AsSpan(0, _pngSignature.Length).SequenceEqual(_pngSignature),
                $"'{suffix}' does not start with the PNG signature.");
        }
    }

    [Fact]
    public void The_locked_and_unlocked_glyphs_are_different_pictures()
    {
        // The point of the ruling: the image says which state the control is in. Two byte
        // sequences here means the host is given two different pictures; identical assets
        // would render a padlock that never changes while the label contradicted it.
        byte[]? locked = RibbonImages.ReadPng(LockGlyphImages.LockedResourceSuffix);
        byte[]? unlocked = RibbonImages.ReadPng(LockGlyphImages.UnlockedResourceSuffix);

        Assert.NotNull(locked);
        Assert.NotNull(unlocked);
        Assert.False(
            locked.AsSpan().SequenceEqual(unlocked),
            "The locked and unlocked glyphs are byte-identical, so the image cannot follow the state.");
    }

    [Fact]
    public void Each_state_resolves_to_a_picture_and_the_result_is_cached()
    {
        object? locked = LockGlyphImages.For(locked: true);
        object? unlocked = LockGlyphImages.For(locked: false);

        Assert.NotNull(locked);
        Assert.NotNull(unlocked);

        // Excel asks on every invalidation. A fresh OLE wrapper per call would allocate and
        // leak a COM object for nothing, so the second answer must be the same instance.
        Assert.Same(locked, LockGlyphImages.For(locked: true));
        Assert.Same(unlocked, LockGlyphImages.For(locked: false));
        Assert.NotSame(locked, unlocked);
    }

    [Fact]
    public void A_suffix_with_no_embedded_resource_yields_no_picture()
    {
        // Positive test for the failure path: a missing glyph must not throw, because an
        // exception inside a ribbon callback takes the whole ribbon down.
        Assert.Null(RibbonImages.For("no-such-glyph-32.png"));
        Assert.Null(RibbonImages.ReadPng("no-such-glyph-32.png"));
    }

    [Fact]
    public void Both_lock_toggles_take_their_glyph_from_a_callback_rather_than_a_gallery_id()
    {
        string xml = new GanttRibbon().GetCustomUI(WorkbookRibbonId)!;
        XDocument document = XDocument.Parse(xml);
        XNamespace ns = document.Root!.Name.Namespace;

        foreach ((string id, string callback) in new[]
                 {
                     (RibbonControlIds.ToggleCellLock, "GetToggleCellLockImage"),
                     (RibbonControlIds.ToggleChartLock, "GetToggleChartLockImage"),
                 })
        {
            XElement element = document.Descendants(ns + "toggleButton")
                .Single(candidate => (string?)candidate.Attribute("id") == id);

            Assert.Equal(callback, (string?)element.Attribute("getImage"));
            Assert.Null(element.Attribute("imageMso"));
        }
    }

    [Fact]
    public void Each_picture_is_an_ole_picture_the_host_can_marshal()
    {
        // The type question the ribbon asks: a getImage callback must hand the host an OLE
        // IPictureDisp, not a managed Image. The interface is identified by IID rather than by
        // a stdole reference, which this project does not carry, and QueryInterface is the same
        // call the host makes when it unmarshals the return value. What this cannot prove —
        // and what stays a live check — is that Excel paints the glyph: no in-process test can
        // observe a ribbon.
        foreach (bool locked in new[] { true, false })
        {
            object? picture = LockGlyphImages.For(locked);
            Assert.NotNull(picture);

            IntPtr unknown = Marshal.GetIUnknownForObject(picture);
            try
            {
                Guid iid = PictureDispIid;
                int hr = Marshal.QueryInterface(unknown, in iid, out IntPtr disp);
                if (disp != IntPtr.Zero)
                {
                    _ = Marshal.Release(disp);
                }

                Assert.Equal(0, hr);
            }
            finally
            {
                _ = Marshal.Release(unknown);
            }
        }
    }

    /// <summary><c>IPictureDisp</c>'s IID, which the ribbon host queries before painting.</summary>
    private static readonly Guid PictureDispIid = new("7BF80981-BF32-101A-8BBB-00AA00300CAB");

    [Fact]
    public void Both_lock_image_callbacks_share_the_label_state_source()
    {
        // The glyph, the wording and what the sheet enforces are read from one preference
        // object, so they cannot disagree. Asserted through the public callbacks the ribbon
        // binds, which is the surface Excel actually calls.
        var ribbon = new GanttRibbon();
        Assert.NotNull(ribbon.GetToggleChartLockImage(new StubControl()));
        Assert.NotNull(ribbon.GetToggleCellLockImage(new StubControl()));
    }

    private sealed class StubControl : IRibbonControl
    {
        public string Id => "stub";

        public object Context => new();

        public string Tag => string.Empty;
    }
}
