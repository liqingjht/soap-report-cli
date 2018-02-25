# soap-report-cli #

<font size=3>Generate excel file from soap result</font>

<font size=2>

Usually we use Python script to test our soap api, then get the result as following:

<pre>
ParentalControl:Authenticate                   PASS  ResponseTime=1186135927   ns

DeviceConfig:CheckAppNewFirmware               : PASS [ResponseTime:3.0458520886    s]
</pre>

Then we should pass out the report to customer, so we should show the result in a better way.

That is what this tool for.


----------


Usage: app [options]


  Options:

    -p, --project <project>  project name. eg: example
    -v, --version <version>  project version. eg: 1.0.2.28
    -s, --spec <spec>        meet SOAP Spec version. eg: 3.18
    -h, --help               output usage information

  Example:

    node app.js -p example -v 1.0.2.28 -s 3.18

  Folders:

    -- ./script_result/2.0_result 3.0_result    [should contain "2.0" and "3.0"]
    -- ./not_support_api/not_support_list       [projectName-specVersion.txt]
    -- ./excel_files/excel_file                 [excel file created by this tool]

</font>