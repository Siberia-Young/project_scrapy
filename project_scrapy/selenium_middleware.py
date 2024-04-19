from scrapy import signals
from scrapy.http import HtmlResponse
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager

class CustomHtmlResponse(HtmlResponse):
    def __init__(self, driver, *args, **kwargs):
        self.webdriver = driver  # 保存 WebDriver 对象到自定义的 HtmlResponse 子类中
        super().__init__(*args, **kwargs)

class SeleniumMiddleware:
    @classmethod
    def from_crawler(cls, crawler):
        middleware = cls()
        crawler.signals.connect(middleware.spider_opened, signal=signals.spider_opened)
        crawler.signals.connect(middleware.spider_closed, signal=signals.spider_closed)
        return middleware

    def spider_opened(self, spider):
        options = webdriver.FirefoxOptions()
        self.driver = webdriver.Remote(command_executor="http://127.0.0.1:4444", options=options)

    def spider_closed(self, spider):
        self.driver.quit()

    def process_request(self, request, spider):
        self.driver.get(request.url)
        body = self.driver.page_source.encode('utf-8')
        return CustomHtmlResponse(self.driver, self.driver.current_url, body=body, encoding='utf-8', request=request)