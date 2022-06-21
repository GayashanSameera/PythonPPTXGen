import imgkit

body = """
<html>
  <head>
    <meta name="imgkit-format" content="png"/>
    <meta name="imgkit-orientation" content="Landscape"/>
  </head>
  Hello World!
</html>
"""

imgkit.from_string(body, 'out.png')
