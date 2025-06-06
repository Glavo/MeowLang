/*
 * Copyright (C) 2025 Glavo
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */
package org.glavo.meow.value;

import org.glavo.meow.Meow;
import org.glavo.meow.MeowContext;
import org.glavo.meow.MeowSymbolMap;
import org.glavo.meow.ast.MeowExpression;

import java.util.List;

@FunctionalInterface
public non-sealed interface MeowMacro extends MeowValue {

    default String getName() {
        return "macro@" + Integer.toHexString(System.identityHashCode(this));
    }

    @Override
    default MeowValue apply(MeowContext context, List<MeowExpression> args) {
        String macro = getName();
        if (Meow.DEBUG) {
            context.log("|   >>> {0}{2} in {1}",
                    macro,
                    context,
                    MeowSymbolMap.INSTANCE.toString(args));
        }
        return process(context, args);
    }

    MeowValue process(MeowContext context, List<MeowExpression> args);
}
