"""Print how many test cases ran (not skipped) in a JUnit report."""

import sys
import xml.etree.ElementTree as ET

root = ET.parse(sys.argv[1]).getroot()
suites = [root] if root.tag == "testsuite" else list(root)
print(sum(int(s.get("tests", 0)) - int(s.get("skipped", 0)) for s in suites))
