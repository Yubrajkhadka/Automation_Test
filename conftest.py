import pytest
import os

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    if report.when == "call" and report.failed:
        driver = item.funcargs.get("setup")
        if driver:
            screenshot_dir = "screenshots"
            os.makedirs(screenshot_dir, exist_ok=True)
            file_name = f"{item.name}.png"
            file_path = os.path.join(screenshot_dir, file_name)
            driver.save_screenshot(file_path)
            if hasattr(item.config, "_html"):
                extra = getattr(report, "extra", [])
                extra.append(pytest_html.extras.png(file_path))
                report.extra = extra
