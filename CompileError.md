<p>The error message “Compile error: Wrong number of arguments or invalid property assignment” in Excel VBA typically occurs when a function or property is being used incorrectly. This can happen when the number of arguments provided to a function does not match the expected number, or when an assignment is made to a property that does not support it.</p>
<p>In the code shown in the screenshot, the error seems to be related to setting the <code>.Formula1</code> property of a <code>FormatCondition</code> object. This property expects a single string that represents the formula to be used for the condition without the leading equals sign (<code>=</code>).</p>
<p>There may be several reasons for this error:</p>
<ol>
<li>
<p><strong>Misconstructed Formula</strong>: The formula string may be incorrectly constructed, with missing or extra quotes, or incorrect concatenation which can cause the VBA compiler to misinterpret the intended arguments.</p>
</li>
<li>
<p><strong>Incorrect Usage of Property</strong>: If the <code>.Formula1</code> property is not expecting a formula in the way it’s being constructed, this could also trigger the error. For instance, if there are syntax errors within the formula string or it’s not a valid Excel formula.</p>
</li>
<li>
<p><strong>Extra Characters</strong>: Sometimes, hidden characters or typos may cause this error. It’s essential to ensure that the formula string is clean and contains no extra characters other than those necessary for the formula itself.</p>
</li>
</ol>
<p>Looking at the line that’s highlighted:</p>
<pre class=" language-vba"><code class="prism  language-vba">fc.Formula1 = "=""AND("" &amp; ColRef &amp; ""&lt;1"",OR("" &amp; ColRef &amp; ""&lt;2"","" $D1&lt;&gt;""))"
</code></pre>
<p>The construction of the formula string seems complex and there might be a mismatch in the quotes or a logical error within the formula construction. It’s also worth checking if <code>ColRef</code> contains the expected string value and if the spaces, particularly before <code>$D1</code>, are intentional and formatted correctly.</p>
<p>To resolve this, ensure that the formula string passed to <code>.Formula1</code> is a valid Excel formula and that it is constructed correctly within the VBA string context. This often involves correctly escaping quotes and concatenating strings in a way that produces a valid result when evaluated by Excel.</p>

