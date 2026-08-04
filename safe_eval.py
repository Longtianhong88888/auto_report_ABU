import ast
import operator

# ================== 安全表达式求值（替代 eval，防止配置文件任意代码执行） ==================
_TRANSFORM_CALLS = {
    'abs': abs, 'round': round, 'min': min, 'max': max,
    'int': int, 'float': float, 'str': str, 'len': len, 'sum': sum,
}
_TRANSFORM_NODES = (
    ast.Expression, ast.Constant, ast.Name, ast.Load,
    ast.BinOp, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.Pow,
    ast.UnaryOp, ast.UAdd, ast.USub,
    ast.Compare, ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
    ast.BoolOp, ast.And, ast.Or, ast.IfExp, ast.Call, ast.keyword,
)
_TRANSFORM_BINOPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod, ast.Pow: operator.pow,
}
_TRANSFORM_UNARYOPS = {ast.UAdd: operator.pos, ast.USub: operator.neg}
_TRANSFORM_CMPOPS = {
    ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt,
    ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge,
}


def _check_transform_expr(expr):
    """校验表达式是否在安全子集内；合法返回 None，否则返回错误描述"""
    try:
        tree = ast.parse(expr.strip(), mode='eval')
    except SyntaxError as e:
        return f"语法错误：{e}"
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            if node.id != 'x' and node.id not in _TRANSFORM_CALLS:
                return f"不支持的变量：{node.id}"
            continue
        if not isinstance(node, _TRANSFORM_NODES):
            return f"不支持的表达式结构：{type(node).__name__}"
        if isinstance(node, ast.Call):
            if not (isinstance(node.func, ast.Name) and node.func.id in _TRANSFORM_CALLS):
                return "不支持的函数调用"
    return None


def _safe_eval_transform(expr, x):
    """在受限 AST 上求值；返回 (是否成功, 结果)"""
    try:
        tree = ast.parse(expr.strip(), mode='eval')
    except SyntaxError:
        return False, None
    try:
        return True, _eval_transform_node(tree.body, x)
    except Exception:
        return False, None


def _eval_transform_node(node, x):
    if isinstance(node, ast.Expression):
        return _eval_transform_node(node.body, x)
    if isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float, bool)):
            return node.value
        raise ValueError("不支持的常量")
    if isinstance(node, ast.Name):
        if node.id == 'x':
            return x
        raise ValueError(f"不支持的变量：{node.id}")
    if isinstance(node, ast.BinOp) and type(node.op) in _TRANSFORM_BINOPS:
        return _TRANSFORM_BINOPS[type(node.op)](
            _eval_transform_node(node.left, x), _eval_transform_node(node.right, x))
    if isinstance(node, ast.UnaryOp) and type(node.op) in _TRANSFORM_UNARYOPS:
        return _TRANSFORM_UNARYOPS[type(node.op)](_eval_transform_node(node.operand, x))
    if isinstance(node, ast.Compare):
        left = _eval_transform_node(node.left, x)
        for op, comp in zip(node.ops, node.comparators):
            right = _eval_transform_node(comp, x)
            if type(op) not in _TRANSFORM_CMPOPS or not _TRANSFORM_CMPOPS[type(op)](left, right):
                return False
            left = right
        return True
    if isinstance(node, ast.BoolOp):
        if isinstance(node.op, ast.And):
            return all(_eval_transform_node(v, x) for v in node.values)
        return any(_eval_transform_node(v, x) for v in node.values)
    if isinstance(node, ast.IfExp):
        return _eval_transform_node(node.body, x) if _eval_transform_node(node.test, x) \
            else _eval_transform_node(node.orelse, x)
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
        fn = _TRANSFORM_CALLS[node.func.id]
        args = [_eval_transform_node(a, x) for a in node.args]
        kwargs = {kw.arg: _eval_transform_node(kw.value, x) for kw in node.keywords if kw.arg}
        return fn(*args, **kwargs)
    raise ValueError("不支持的表达式")

